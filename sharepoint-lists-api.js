// ═══════════════════════════════════════════════════════════════════════════
// SharePointListsAPI v3 — Module centralisé Microsoft Lists via Graph API
// Audit 5S SCE — Sartorius Pompey
// ═══════════════════════════════════════════════════════════════════════════
//
// Colonnes person : Equipe5S.Nom, Audits5S.Auditeur, Actions5S.Responsable
//   → Lecture : Graph renvoie {LookupId, LookupValue, Email}
//   → Écriture : envoyer {colonneLookupId: <id numérique>}
//   → On cache la correspondance nom → LookupId via siteUsers
//
// Accents encodés dans les noms internes SharePoint :
//   é → _x00e9_   è → _x00e8_   ô → _x00f4_
//
// Dépendance : MSAL.js, msalInstance + msalAccount + loginRequest (globales)
// Scopes : Sites.ReadWrite.All
// ═══════════════════════════════════════════════════════════════════════════

const SharePointListsAPI = (() => {
    'use strict';

    // ── Configuration ──────────────────────────────────────────────────────
    const config = {
        siteId: 'sartorius.sharepoint.com,ec2b83d2-7066-467c-83e6-fa97ca5db8e3,a833e672-c3d7-4db5-92c0-6169773c2516',

        lists: {
            equipe:  'a21b962a-0987-48b9-a519-9e30b98efbec',
            audits:  'fb669de2-4d64-49c7-a1cb-50a3b7ed79dc',
            actions: 'aa10c625-373c-4d14-859a-3b468470cc3f'
        },

        listNames: {
            equipe:  'Equipe5S',
            audits:  'Audits5S',
            actions: 'Actions5S'
        },

        maxRetries: 3,
        retryDelayMs: 1000,
        requestTimeoutMs: 30000
    };

    // ═══════════════════════════════════════════════════════════════════════
    //  MAPPING NOMS DE COLONNES — logique → nom interne SharePoint
    //
    //  ┌─────────────────────────────────────────────────────────────┐
    //  │ Equipe5S                                                    │
    //  │   Title (renommé "Trigramme") = trigramme (GDE, HBI...)    │
    //  │   Nom = colonne person (photo + validation)                 │
    //  │   Email, ZoneResponsable, EstAuditeur, Actif, Niveau       │
    //  ├─────────────────────────────────────────────────────────────┤
    //  │ Audits5S                                                    │
    //  │   Title = AuditID                                           │
    //  │   Auditeur = colonne person                                 │
    //  │   Niveau = choice ("1", "2", "3")                           │
    //  ├─────────────────────────────────────────────────────────────┤
    //  │ Actions5S                                                   │
    //  │   Title = ActionID                                          │
    //  │   Responsable = colonne person                              │
    //  │   TypeAction = choice (pas "Type" car réservé)              │
    //  │   Accents : DateCréation, Critère, DateClôture,             │
    //  │             CommentaireClôture                               │
    //  │   Photos = pièces jointes (pas colonnes image)              │
    //  └─────────────────────────────────────────────────────────────┘
    // ═══════════════════════════════════════════════════════════════════════

    const COLUMN_MAP = {
        equipe: {
            trigramme:       'Title',               // Title renommé "Trigramme"
            nom:             'Nom',                  // Person column
            nomLookupId:     'NomLookupId',          // Pour écriture person
            email:           'Email',
            zoneResponsable: 'ZoneResponsable',
            estAuditeur:     'EstAuditeur',
            actif:           'Actif',
            niveau:          'Niveau'                // Choice — niveau par défaut de l'auditeur
        },

        audits: {
            auditId:         'Title',
            date:            'Date',
            auditeur:        'Auditeur',             // Person column (lecture)
            auditeurLookupId:'AuditeurLookupId',     // Person column (écriture)
            auditeurEmail:   'AuditeurEmail',        // Texte — fallback/complément
            zone:            'Zone',
            niveau:          'Niveau',               // Choice → "1", "2", "3"
            s1_1: 'S1_1', s1_2: 'S1_2', s1_3: 'S1_3',
            s2_1: 'S2_1', s2_2: 'S2_2', s2_3: 'S2_3',
            s3_1: 'S3_1', s3_2: 'S3_2', s3_3: 'S3_3',
            s4_1: 'S4_1', s4_2: 'S4_2', s4_3: 'S4_3',
            s5_1: 'S5_1', s5_2: 'S5_2', s5_3: 'S5_3',
            scoreTotal:      'ScoreTotal',
            scorePct:        'ScorePct',
            commentaire:     'Commentaire'
        },

        actions: {
            actionId:            'Title',
            auditRef:            'AuditRefLookupId',
            dateCreation:        'Date',                    // "DateCréation" — nom interne = Date
            zone:                'Zone',
            critere:             'Crit_x00e8_re',          // "Critère"
            type:                'TypeAction',              // "Type" était réservé
            description:         'Description',
            responsable:         'Responsable',             // Person column (lecture)
            responsableLookupId: 'ResponsableLookupId',     // Person column (écriture)
            statut:              'Statut',
            dateCloture:         'DateCloture',              // "DateClôture" — recréée sans accent
            commentaireCloture:  'CommentaireCloture'        // "CommentaireClôture" — recréée sans accent
        }
    };

    // ── Cache ──────────────────────────────────────────────────────────────
    let _equipeCache = null;
    let _equipeCacheTimestamp = 0;
    const EQUIPE_CACHE_TTL = 5 * 60 * 1000;

    // Cache siteUsers : email → LookupId (pour colonnes person)
    let _siteUsersCache = null;

    let _columnNamesCache = {};

    // ── Sécurité ───────────────────────────────────────────────────────────
    function escapeHtml(str) {
        if (typeof str !== 'string') return '';
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    // ── Token ──────────────────────────────────────────────────────────────
    async function getToken() {
        if (!msalInstance || !msalAccount) {
            throw new Error('SharePointListsAPI: Non authentifié');
        }
        try {
            const r = await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account: msalAccount });
            return r.accessToken;
        } catch (e) {
            if (typeof isMobileDevice === 'function' && isMobileDevice()) {
                await msalInstance.loginRedirect(loginRequest);
                return null;
            }
            const r = await msalInstance.loginPopup(loginRequest);
            msalAccount = r.account;
            const t = await msalInstance.acquireTokenSilent({ scopes: loginRequest.scopes, account: msalAccount });
            return t.accessToken;
        }
    }

    // ── Fetch avec retry/timeout/throttle ──────────────────────────────────
    async function graphFetch(url, options = {}, retryCount = 0) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), config.requestTimeoutMs);

        try {
            const token = await getToken();
            const response = await fetch(url, {
                ...options,
                signal: controller.signal,
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json',
                    ...(options.headers || {})
                }
            });
            clearTimeout(timeoutId);

            if (response.status === 429 && retryCount < config.maxRetries) {
                const wait = parseInt(response.headers.get('Retry-After') || '5', 10);
                await sleep(wait * 1000);
                return graphFetch(url, options, retryCount + 1);
            }
            if (response.status >= 500 && retryCount < config.maxRetries) {
                await sleep(config.retryDelayMs * (retryCount + 1));
                return graphFetch(url, options, retryCount + 1);
            }
            if (!response.ok) {
                const err = await response.json().catch(() => ({}));
                const msg = err?.error?.message || `HTTP ${response.status}`;
                console.error('SharePointListsAPI:', { url: url.substring(0, 150), status: response.status, error: err?.error });
                throw new Error(`Graph API: ${msg} (${response.status})`);
            }
            if (response.status === 204) return null;
            return await response.json();
        } catch (error) {
            clearTimeout(timeoutId);
            if (error.name === 'AbortError') {
                if (retryCount < config.maxRetries) return graphFetch(url, options, retryCount + 1);
                throw new Error('Graph API: Timeout');
            }
            throw error;
        }
    }

    function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

    // ── URL helpers ────────────────────────────────────────────────────────
    function itemsUrl(listKey) {
        return `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists[listKey]}/items`;
    }

    async function resolveListIds() {
        if (config.lists.equipe && config.lists.audits && config.lists.actions) return;
        const base = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists`;
        for (const [key, name] of Object.entries(config.listNames)) {
            if (config.lists[key]) continue;
            const data = await graphFetch(`${base}?$filter=displayName eq '${name}'&$select=id`);
            if (data.value?.length > 0) config.lists[key] = data.value[0].id;
            else throw new Error(`Liste "${name}" introuvable`);
        }
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  COLONNES PERSON — Résolution LookupId via siteUsers
    //
    //  Les colonnes person de SharePoint stockent un LookupId qui pointe
    //  vers la liste cachée "User Information List" du site.
    //  Pour écrire dans une colonne person, on a besoin de ce LookupId.
    //
    //  Stratégie :
    //  1. Charger les siteUsers une fois (GET /sites/{id}/lists/User%20Information%20List/items)
    //     ou via l'API simplifiée ensureUser
    //  2. Cacher email → LookupId
    //  3. En écriture : résoudre le LookupId depuis le cache
    //  4. En lecture : Graph renvoie le LookupValue (nom affiché)
    // ═══════════════════════════════════════════════════════════════════════

    /**
     * Charge et cache la correspondance email → site user LookupId.
     * Utilise l'endpoint siteUsers de Graph.
     */
    async function loadSiteUsers() {
        if (_siteUsersCache) return _siteUsersCache;

        // Méthode : GET /sites/{siteId}/lists('User Information List')/items
        // Cette liste système contient tous les utilisateurs connus du site
        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists('User%20Information%20List')/items?$select=id&$expand=fields($select=EMail,Title,Name)&$top=500`;

        try {
            const data = await graphFetch(url);
            _siteUsersCache = new Map();

            (data.value || []).forEach(item => {
                const f = item.fields;
                const email = (f.EMail || '').toLowerCase();
                const name = f.Title || '';
                if (email) {
                    _siteUsersCache.set(email, { lookupId: parseInt(item.id), name, email: f.EMail });
                }
                // Aussi indexer par nom pour le fallback
                if (name) {
                    _siteUsersCache.set('name:' + name.toLowerCase(), { lookupId: parseInt(item.id), name, email: f.EMail });
                }
            });

            console.log(`SharePointListsAPI: ${_siteUsersCache.size / 2} utilisateurs site chargés`);
            return _siteUsersCache;
        } catch (e) {
            console.warn('SharePointListsAPI: Impossible de charger siteUsers —', e.message);
            console.warn('Les colonnes person seront laissées vides en écriture.');
            _siteUsersCache = new Map();
            return _siteUsersCache;
        }
    }

    /**
     * Résout un email ou un nom en LookupId pour une colonne person.
     * @param {string} emailOrName — email ou nom complet
     * @returns {number|null} — LookupId ou null si non trouvé
     */
    async function resolvePersonLookupId(emailOrName) {
        if (!emailOrName) return null;
        const users = await loadSiteUsers();

        // Essayer par email d'abord
        const byEmail = users.get(emailOrName.toLowerCase());
        if (byEmail) return byEmail.lookupId;

        // Puis par nom
        const byName = users.get('name:' + emailOrName.toLowerCase());
        if (byName) return byName.lookupId;

        // Fallback : essayer ensureUser via API (créé l'entrée si nécessaire)
        // Note : cet endpoint n'est pas disponible via Graph standard,
        // on le tente via le REST SharePoint classique
        console.warn(`SharePointListsAPI: Utilisateur "${emailOrName}" non trouvé dans siteUsers`);
        return null;
    }

    /**
     * Lit la valeur d'une colonne person depuis un item Graph.
     * @param {Object} fields — les fields d'un item Graph
     * @param {string} colName — nom de la colonne person (ex: 'Auditeur')
     * @returns {Object} — {lookupId, displayName, email}
     */
    function readPersonField(fields, colName) {
        // Graph renvoie la colonne person comme un objet avec LookupId + LookupValue
        // Mais via $expand=fields, le format peut varier.
        // Le LookupId est dans fields[colName + 'LookupId']
        // Le nom affiché est dans fields[colName] (si c'est un string) ou fields[colName].LookupValue
        const lookupId = fields[colName + 'LookupId'];
        const rawValue = fields[colName];

        let displayName = '';
        if (typeof rawValue === 'string') {
            displayName = rawValue;
        } else if (rawValue && rawValue.LookupValue) {
            displayName = rawValue.LookupValue;
        }

        return {
            lookupId: lookupId || null,
            displayName: displayName,
            email: '' // L'email n'est pas dans le champ person directement
        };
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  EQUIPE5S
    // ═══════════════════════════════════════════════════════════════════════

    async function loadEquipe(forceRefresh = false) {
        const now = Date.now();
        if (!forceRefresh && _equipeCache && (now - _equipeCacheTimestamp < EQUIPE_CACHE_TTL)) {
            return _equipeCache;
        }

        await resolveListIds();
        const C = COLUMN_MAP.equipe;

        // On demande tous les champs pertinents y compris la colonne person
        const selectFields = [
            C.trigramme, C.email, C.zoneResponsable,
            C.estAuditeur, C.actif, C.niveau,
            C.nom, C.nomLookupId  // Person : nom + lookupId
        ].filter(Boolean).join(',');

        const url = `${itemsUrl('equipe')}?$expand=fields($select=${selectFields})&$top=200`;
        const data = await graphFetch(url);

        _equipeCache = (data.value || [])
            .filter(item => item.fields)
            .map(item => {
                const f = item.fields;
                const person = readPersonField(f, C.nom);

                return {
                    id:              item.id,
                    trigramme:       f[C.trigramme] || '',       // Title = trigramme
                    nom:             person.displayName || '',    // Colonne person
                    nomLookupId:     person.lookupId || null,
                    email:           f[C.email] || '',
                    zoneResponsable: f[C.zoneResponsable] || null,
                    estAuditeur:     f[C.estAuditeur] === true,
                    actif:           f[C.actif] !== false,
                    niveau:          f[C.niveau] || null          // Niveau par défaut
                };
            });

        _equipeCacheTimestamp = now;
        console.log(`SharePointListsAPI: Equipe chargée — ${_equipeCache.length} personnes`);
        return _equipeCache;
    }

    async function getActiveMembers() {
        return (await loadEquipe()).filter(p => p.actif);
    }

    async function getAuditeurs() {
        return (await loadEquipe()).filter(p => p.actif && p.estAuditeur);
    }

    async function getEmailByName(name) {
        const p = (await loadEquipe()).find(p => p.nom === name);
        return p ? p.email : '';
    }

    async function isValidName(name) {
        return (await loadEquipe()).some(p => p.nom === name && p.actif);
    }

    async function getTrigrammeByName(name) {
        const p = (await loadEquipe()).find(p => p.nom === name);
        return p ? p.trigramme : '';
    }

    /**
     * Résout le LookupId person pour un nom d'équipier.
     * Cherche d'abord dans le cache Equipe (qui a le NomLookupId),
     * puis fallback vers siteUsers.
     */
    async function resolveEquipierLookupId(name) {
        // 1. Chercher dans Equipe5S (déjà lu)
        const equipe = await loadEquipe();
        const member = equipe.find(p => p.nom === name);
        if (member?.nomLookupId) return member.nomLookupId;

        // 2. Chercher par email dans siteUsers
        if (member?.email) {
            const id = await resolvePersonLookupId(member.email);
            if (id) return id;
        }

        // 3. Chercher par nom dans siteUsers
        return await resolvePersonLookupId(name);
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  AUDITS5S
    // ═══════════════════════════════════════════════════════════════════════

    async function createAudit(audit) {
        await resolveListIds();
        const C = COLUMN_MAP.audits;

        const fields = {
            [C.auditId]:       audit.auditId,
            [C.date]:          audit.date,
            [C.auditeurEmail]: audit.auditeurEmail || audit.auditeur || '',
            [C.zone]:          audit.zone,
            [C.niveau]:        String(audit.niveau),
            [C.s1_1]: audit.scores.seiri_1    ?? 0,
            [C.s1_2]: audit.scores.seiri_2    ?? 0,
            [C.s1_3]: audit.scores.seiri_3    ?? 0,
            [C.s2_1]: audit.scores.seiton_1   ?? 0,
            [C.s2_2]: audit.scores.seiton_2   ?? 0,
            [C.s2_3]: audit.scores.seiton_3   ?? 0,
            [C.s3_1]: audit.scores.seiso_1    ?? 0,
            [C.s3_2]: audit.scores.seiso_2    ?? 0,
            [C.s3_3]: audit.scores.seiso_3    ?? 0,
            [C.s4_1]: audit.scores.seiketsu_1 ?? 0,
            [C.s4_2]: audit.scores.seiketsu_2 ?? 0,
            [C.s4_3]: audit.scores.seiketsu_3 ?? 0,
            [C.s5_1]: audit.scores.shitsuke_1 ?? 0,
            [C.s5_2]: audit.scores.shitsuke_2 ?? 0,
            [C.s5_3]: audit.scores.shitsuke_3 ?? 0,
            [C.scoreTotal]:    audit.scoreTotal,
            [C.scorePct]:      audit.scorePct,
            [C.commentaire]:   audit.commentaire || ''
        };

        // Colonne person "Auditeur" : résoudre le LookupId
        // On cherche par email d'abord (plus fiable), puis par nom
        if (audit.auditeurEmail || audit.auditeur) {
            let lookupId = null;
            // 1. Par email direct (le plus fiable)
            if (audit.auditeurEmail && audit.auditeurEmail.includes('@')) {
                lookupId = await resolvePersonLookupId(audit.auditeurEmail);
            }
            // 2. Par email depuis Equipe5S
            if (!lookupId && audit.auditeur) {
                const email = await getEmailByName(audit.auditeur);
                if (email) lookupId = await resolvePersonLookupId(email);
            }
            // 3. Par nom dans siteUsers (dernier recours)
            if (!lookupId && audit.auditeur) {
                lookupId = await resolvePersonLookupId(audit.auditeur);
            }
            if (lookupId) {
                fields[C.auditeurLookupId] = lookupId;
            } else {
                console.warn(`SharePointListsAPI: LookupId non trouvé pour "${audit.auditeur}" (${audit.auditeurEmail}) — colonne person laissée vide`);
            }
        }

        const result = await graphFetch(itemsUrl('audits'), {
            method: 'POST',
            body: JSON.stringify({ fields })
        });

        console.log(`SharePointListsAPI: Audit créé — ${audit.auditId} (#${result.id})`);
        return result;
    }

    async function getAudits(options = {}) {
        await resolveListIds();
        const C = COLUMN_MAP.audits;

        const selectFields = [
            C.auditId, C.date, C.auditeurEmail, C.zone, C.niveau,
            C.auditeur, C.auditeurLookupId,  // Person field
            C.s1_1, C.s1_2, C.s1_3, C.s2_1, C.s2_2, C.s2_3,
            C.s3_1, C.s3_2, C.s3_3, C.s4_1, C.s4_2, C.s4_3,
            C.s5_1, C.s5_2, C.s5_3, C.scoreTotal, C.scorePct, C.commentaire
        ].join(',');

        const filters = [];
        if (options.zone) filters.push(`fields/${C.zone} eq '${options.zone}'`);
        if (options.year) {
            filters.push(`fields/${C.date} ge '${options.year}-01-01'`);
            filters.push(`fields/${C.date} le '${options.year}-12-31'`);
        }

        let url = `${itemsUrl('audits')}?$expand=fields($select=${selectFields})&$top=${options.top || 500}`;
        if (filters.length > 0) url += `&$filter=${filters.join(' and ')}`;

        const data = await graphFetch(url);

        const audits = (data.value || []).map(item => {
            const f = item.fields;
            const auditeurPerson = readPersonField(f, C.auditeur);

            // Nom de l'auditeur : person > AuditeurEmail > fallback
            const auditeurNom = auditeurPerson.displayName || f[C.auditeurEmail] || '';

            return {
                id:            item.id,
                auditId:       f[C.auditId],
                date:          f[C.date],
                auditeur:      auditeurNom,
                auditeurEmail: f[C.auditeurEmail] || '',
                zone:          f[C.zone],
                niveau:        parseInt(f[C.niveau]) || 1,
                scores: {
                    seiri_1: f[C.s1_1]??0, seiri_2: f[C.s1_2]??0, seiri_3: f[C.s1_3]??0,
                    seiton_1: f[C.s2_1]??0, seiton_2: f[C.s2_2]??0, seiton_3: f[C.s2_3]??0,
                    seiso_1: f[C.s3_1]??0, seiso_2: f[C.s3_2]??0, seiso_3: f[C.s3_3]??0,
                    seiketsu_1: f[C.s4_1]??0, seiketsu_2: f[C.s4_2]??0, seiketsu_3: f[C.s4_3]??0,
                    shitsuke_1: f[C.s5_1]??0, shitsuke_2: f[C.s5_2]??0, shitsuke_3: f[C.s5_3]??0
                },
                scoreS1: ((f[C.s1_1]??0)+(f[C.s1_2]??0)+(f[C.s1_3]??0))/3,
                scoreS2: ((f[C.s2_1]??0)+(f[C.s2_2]??0)+(f[C.s2_3]??0))/3,
                scoreS3: ((f[C.s3_1]??0)+(f[C.s3_2]??0)+(f[C.s3_3]??0))/3,
                scoreS4: ((f[C.s4_1]??0)+(f[C.s4_2]??0)+(f[C.s4_3]??0))/3,
                scoreS5: ((f[C.s5_1]??0)+(f[C.s5_2]??0)+(f[C.s5_3]??0))/3,
                scoreTotal: f[C.scoreTotal] ?? 0,
                scorePct:   f[C.scorePct] ?? 0,
                commentaire: f[C.commentaire] || ''
            };
        });

        audits.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
        return audits;
    }

    async function getAuditById(itemId) {
        return await graphFetch(`${itemsUrl('audits')}/${itemId}?$expand=fields`);
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  ACTIONS5S
    // ═══════════════════════════════════════════════════════════════════════

    async function createAction(action) {
        await resolveListIds();
        const C = COLUMN_MAP.actions;

        const fields = {
            [C.actionId]:     action.actionId,
            [C.dateCreation]: action.dateCreation || new Date().toISOString().split('T')[0],
            [C.zone]:         action.zone,
            [C.critere]:      action.critere,
            [C.type]:         action.type,
            [C.description]:  action.description,
            [C.statut]:       'Ouverte'
        };

        // Lookup vers Audits5S
        if (action.auditItemId) {
            fields[C.auditRef] = parseInt(action.auditItemId);
        }

        // Colonne person "Responsable" : résoudre le LookupId
        if (action.responsable || action.responsableEmail) {
            let lookupId = null;
            // 1. Par email direct
            if (action.responsableEmail && action.responsableEmail.includes('@')) {
                lookupId = await resolvePersonLookupId(action.responsableEmail);
            }
            // 2. Par email depuis Equipe5S
            if (!lookupId && action.responsable) {
                const email = await getEmailByName(action.responsable);
                if (email) lookupId = await resolvePersonLookupId(email);
            }
            // 3. Par nom
            if (!lookupId && action.responsable) {
                lookupId = await resolvePersonLookupId(action.responsable);
            }
            if (lookupId) {
                fields[C.responsableLookupId] = lookupId;
            } else {
                console.warn(`SharePointListsAPI: LookupId non trouvé pour responsable "${action.responsable}"`);
            }
        }

        const result = await graphFetch(itemsUrl('actions'), {
            method: 'POST',
            body: JSON.stringify({ fields })
        });

        console.log(`SharePointListsAPI: Action créée — ${action.actionId} (#${result.id})`);
        return result;
    }

    async function getActions(options = {}) {
        await resolveListIds();
        const C = COLUMN_MAP.actions;

        const selectFields = [
            C.actionId, C.dateCreation, C.zone, C.critere,
            C.type, C.description,
            C.responsable, C.responsableLookupId,  // Person field
            C.statut, C.dateCloture, C.commentaireCloture
        ].filter(Boolean).join(',');

        const filters = [];
        if (options.zone) filters.push(`fields/${C.zone} eq '${options.zone}'`);
        if (options.statut) filters.push(`fields/${C.statut} eq '${options.statut}'`);

        let url = `${itemsUrl('actions')}?$expand=fields($select=${selectFields})&$top=${options.top || 500}`;
        if (filters.length > 0) url += `&$filter=${filters.join(' and ')}`;

        const data = await graphFetch(url);

        const actions = (data.value || []).map(item => {
            const f = item.fields;
            const respPerson = readPersonField(f, C.responsable);

            // Email du responsable : lookup dans Equipe5S via le nom
            const respNom = respPerson.displayName || '';

            return {
                id:                 item.id,
                actionId:           f[C.actionId],
                auditRefId:         f[C.auditRef],
                dateCreation:       f[C.dateCreation] || '',
                zone:               f[C.zone],
                critere:            f[C.critere] || '',
                type:               f[C.type] || '',
                description:        f[C.description] || '',
                action:             f[C.description] || '',
                responsable:        respNom,
                responsableEmail:   '', // Résolu à la volée si besoin via getEmailByName
                statut:             f[C.statut] || 'Ouverte',
                dateCloture:        f[C.dateCloture] || '',
                commentaireCloture: f[C.commentaireCloture] || '',
                dateOuverture:      f[C.dateCreation] || '',
                commentaire:        f[C.commentaireCloture] || ''
            };
        });

        actions.sort((a, b) => (b.dateCreation || '').localeCompare(a.dateCreation || ''));
        return actions;
    }

    async function updateAction(itemId, updates) {
        await resolveListIds();
        const C = COLUMN_MAP.actions;

        const keyMap = {
            'Statut':              C.statut,
            'DateCloture':         C.dateCloture,
            'CommentaireCloture':  C.commentaireCloture,
            'Description':         C.description,
            'Zone':                C.zone,
            'Type':                C.type,
            'Critere':             C.critere,
            'DateCreation':        C.dateCreation
        };

        const translated = {};
        for (const [key, value] of Object.entries(updates)) {
            translated[keyMap[key] || key] = value;
        }

        const result = await graphFetch(`${itemsUrl('actions')}/${itemId}/fields`, {
            method: 'PATCH',
            body: JSON.stringify(translated)
        });

        console.log(`SharePointListsAPI: Action #${itemId} mise à jour`, Object.keys(updates));
        return result;
    }

    async function closeAction(itemId, comment = '') {
        return updateAction(itemId, {
            Statut: 'Clôturée',
            DateCloture: new Date().toISOString().split('T')[0],
            CommentaireCloture: comment
        });
    }

    async function cancelAction(itemId, comment = '') {
        return updateAction(itemId, {
            Statut: 'Annulée',
            CommentaireCloture: comment
        });
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  PHOTOS — Pièces jointes (attachments) sur items de liste
    // ═══════════════════════════════════════════════════════════════════════

    /**
     * Ajoute une photo comme pièce jointe à un item de liste.
     * Utilise l'API SharePoint REST pour les attachments
     * (Graph API n'a pas d'endpoint natif pour list item attachments,
     * on passe par le endpoint SharePoint classique via Graph proxy).
     */
    async function addAttachment(listKey, itemId, dataUrl, filename) {
        await resolveListIds();

        // Extraire le base64 et convertir en ArrayBuffer
        const base64Data = dataUrl.split(',')[1];
        const byteChars = atob(base64Data);
        const byteArray = new Uint8Array(byteChars.length);
        for (let i = 0; i < byteChars.length; i++) byteArray[i] = byteChars.charCodeAt(i);
        let blob = new Blob([byteArray], { type: 'image/jpeg' });

        // Compresser si > 100KB
        if (blob.size > 100 * 1024) {
            try { blob = await compressImage(blob, 600, 0.7); } catch(e) {}
        }

        // Upload via le document library associé aux attachments de la liste
        // Les pièces jointes de liste sont stockées dans :
        //   /sites/{site}/Lists/{listName}/Attachments/{itemId}/{filename}
        const token = await getToken();
        const listName = config.listNames[listKey];
        const folderPath = `/Lists/${listName}/Attachments/${itemId}`;

        // D'abord tenter de créer le dossier d'attachments (peut déjà exister)
        const uploadUrl = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/drive/root:${folderPath}/${encodeURIComponent(filename)}:/content`;

        const response = await fetch(uploadUrl, {
            method: 'PUT',
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'image/jpeg'
            },
            body: blob
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(`Upload attachment échoué: ${err?.error?.message || response.status}`);
        }

        console.log(`SharePointListsAPI: Attachment "${filename}" → ${listName} #${itemId}`);
        return await response.json();
    }

    /**
     * Récupère la liste des pièces jointes d'un item.
     */
    async function getAttachments(listKey, itemId) {
        const token = await getToken();
        const listName = config.listNames[listKey];
        const folderPath = `/Lists/${listName}/Attachments/${itemId}`;

        try {
            const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/drive/root:${folderPath}:/children?$select=id,name,size,webUrl,@microsoft.graph.downloadUrl`;
            const data = await graphFetch(url);

            return (data.value || []).map(file => ({
                name:       file.name,
                size:       file.size,
                webUrl:     file.webUrl,
                downloadUrl: file['@microsoft.graph.downloadUrl']
            }));
        } catch (e) {
            // Pas de dossier d'attachments = pas de pièces jointes
            return [];
        }
    }

    function compressImage(blob, maxWidth = 600, quality = 0.7) {
        return new Promise((resolve, reject) => {
            const img = new Image();
            img.onload = () => {
                const ratio = Math.min(maxWidth / img.width, 1);
                const canvas = document.createElement('canvas');
                canvas.width = img.width * ratio;
                canvas.height = img.height * ratio;
                canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
                canvas.toBlob(b => b ? resolve(b) : reject(new Error('toBlob null')), 'image/jpeg', quality);
            };
            img.onerror = reject;
            img.src = URL.createObjectURL(blob);
        });
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  SOUMISSION COMPLÈTE
    // ═══════════════════════════════════════════════════════════════════════

    async function submitFullAudit(payload) {
        const { auditId, date, auditeur, auditeurEmail, zone, niveau,
                scores, scoreTotal, scorePct, commentaire, actions } = payload;

        // Pré-charger siteUsers en parallèle avec l'équipe
        await Promise.all([loadEquipe(), loadSiteUsers()]);

        // 1. Créer l'audit
        const auditItem = await createAudit({
            auditId, date, auditeur, auditeurEmail, zone, niveau,
            scores, scoreTotal, scorePct, commentaire
        });

        // 2. Créer les actions par batch de 3
        const actionEntries = Object.entries(actions || {})
            .filter(([, d]) => d.action && d.responsable);

        const actionItems = [];
        for (let i = 0; i < actionEntries.length; i += 3) {
            const batch = actionEntries.slice(i, i + 3);
            const results = await Promise.all(batch.map(async ([critere, d]) => {
                const result = await createAction({
                    actionId: `ACT-${date}-${critere}-${zone.toLowerCase().replace(/[^a-z]/g, '')}`,
                    auditItemId: auditItem.id,
                    dateCreation: date,
                    zone, critere,
                    type: d.type === 'ecart' ? 'Écart aux standards' : 'Élévation des standards',
                    description: d.action,
                    responsable: d.responsable,
                    responsableEmail: d.responsableEmail || await getEmailByName(d.responsable)
                });

                // 3. Photos en pièces jointes
                if (d.photos?.length > 0) {
                    for (let j = 0; j < d.photos.length; j++) {
                        try {
                            await addAttachment('actions', result.id, d.photos[j], `photo_${critere}_${j+1}.jpg`);
                        } catch(e) { console.error(`Photo ${critere}_${j+1}:`, e); }
                    }
                }
                return result;
            }));
            actionItems.push(...results);
        }

        console.log(`SharePointListsAPI: Audit complet — ${auditId}, ${actionItems.length} actions`);
        return { auditItem, actionItems };
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  DIAGNOSTIC
    // ═══════════════════════════════════════════════════════════════════════

    async function discoverColumns(listKey) {
        if (_columnNamesCache[listKey]) return _columnNamesCache[listKey];
        const url = `https://graph.microsoft.com/v1.0/sites/${config.siteId}/lists/${config.lists[listKey]}/columns?$top=100&$select=name,displayName,readOnly,hidden`;
        const data = await graphFetch(url);
        const cols = (data.value || []).filter(c => !c.readOnly && !c.hidden).map(c => ({ name: c.name, displayName: c.displayName }));
        _columnNamesCache[listKey] = cols;
        return cols;
    }

    async function diagnose() {
        const results = { ok: true, checks: [] };

        try { await getToken(); results.checks.push({ name: 'Token', ok: true }); }
        catch(e) { return { ok: false, checks: [{ name: 'Token', ok: false, detail: e.message }] }; }

        try {
            const site = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${config.siteId}`);
            results.checks.push({ name: 'Site', ok: true, detail: site.displayName });
        } catch(e) { results.ok = false; results.checks.push({ name: 'Site', ok: false, detail: e.message }); }

        // Listes
        for (const [key, name] of Object.entries(config.listNames)) {
            try {
                await graphFetch(`${itemsUrl(key)}?$top=1`);
                results.checks.push({ name: `Liste ${name}`, ok: true });
            } catch(e) { results.ok = false; results.checks.push({ name: `Liste ${name}`, ok: false, detail: e.message }); }

            try {
                const cols = await discoverColumns(key);
                const colNames = new Set(cols.map(c => c.name));
                for (const [logical, internal] of Object.entries(COLUMN_MAP[key])) {
                    if (!colNames.has(internal)) {
                        results.ok = false;
                        results.checks.push({ name: `${name}.${logical}`, ok: false, detail: `"${internal}" introuvable` });
                    }
                }
            } catch(e) { results.checks.push({ name: `Colonnes ${name}`, ok: false, detail: e.message }); }
        }

        // siteUsers
        try {
            const users = await loadSiteUsers();
            results.checks.push({ name: 'siteUsers', ok: true, detail: `${users.size / 2} utilisateurs` });
        } catch(e) { results.checks.push({ name: 'siteUsers', ok: false, detail: e.message }); }

        return results;
    }

    // ═══════════════════════════════════════════════════════════════════════
    //  API PUBLIQUE
    // ═══════════════════════════════════════════════════════════════════════

    return {
        config, COLUMN_MAP,
        resolveListIds, discoverColumns, diagnose,

        // Equipe
        loadEquipe, getActiveMembers, getAuditeurs,
        getEmailByName, isValidName, getTrigrammeByName,
        resolveEquipierLookupId,

        // Person columns
        loadSiteUsers, resolvePersonLookupId, readPersonField,

        // Audits
        createAudit, getAudits, getAuditById,

        // Actions
        createAction, getActions, updateAction, closeAction, cancelAction,

        // Photos (attachments)
        addAttachment, getAttachments, compressImage,

        // Soumission complète
        submitFullAudit,

        // Utilitaires
        escapeHtml, graphFetch
    };
})();
