// ---------- Core selectors ----------
    const monthEl = document.getElementById('month');
    const appNav = document.getElementById('appNav');
    const menuToggle = document.getElementById('menuToggle');
    const mobileMenuSignInBtn = document.getElementById('mobileMenuSignInBtn');
    const mobileMenuSignOutBtn = document.getElementById('mobileMenuSignOutBtn');
    const homeMonthLabel = document.getElementById('homeMonthLabel');
    const homeMonthlyTotal = document.getElementById('homeMonthlyTotal');
    const homeEntryCount = document.getElementById('homeEntryCount');
    const modeBadge = document.getElementById('modeBadge');
    const lastEditedEl = document.getElementById('lastEdited');
    const cloudSyncStatus = document.getElementById('cloudSyncStatus');
    const saveCloudBtn = document.getElementById('saveCloudBtn');
    const retryCloudBtn = document.getElementById('retryCloudBtn');
    const exportDataBtn = document.getElementById('exportDataBtn');
    const cloudStatusDetail = document.getElementById('cloudStatusDetail');

    const personNameEl = document.getElementById('personName');
    const addPersonBtn = document.getElementById('addPersonBtn');
    const peopleChips = document.getElementById('peopleChips');

    const categoryNameEl = document.getElementById('categoryName');
    const addCategoryBtn = document.getElementById('addCategoryBtn');
    const categoriesChips = document.getElementById('categoriesChips');

    const modeQuickBtn = document.getElementById('modeQuick');
    const modeBatchBtn = document.getElementById('modeBatch');
    const modeImportBtn = document.getElementById('modeImport');
    const quickPanel = document.getElementById('quickPanel');
    const batchPanel = document.getElementById('batchPanel');
    const importPanel = document.getElementById('importPanel');

    const quickPerson = document.getElementById('quickPerson');
    const quickCategory = document.getElementById('quickCategory');
    const quickSplitSliders = document.getElementById('quickSplitSliders');
    const quickAmounts = document.getElementById('quickAmounts');
    const quickDate = document.getElementById('quickDate');
    const addQuickBtn = document.getElementById('addQuickBtn');
    const clearQuickBtn = document.getElementById('clearQuickBtn');

    const batchPerson = document.getElementById('batchPerson');
    const batchSplitSliders = document.getElementById('batchSplitSliders');
    const batchDate = document.getElementById('batchDate');
    const batchRows = document.getElementById('batchRows');
    const saveBatchBtn = document.getElementById('saveBatchBtn');

    const expensesEmpty = document.getElementById('expensesEmpty');
    const expensesTable = document.getElementById('expensesTable');
    const expensesBody = document.getElementById('expensesBody');
    const expensePayerTabs = document.getElementById('expensePayerTabs');
    const duplicateReview = document.getElementById('duplicateReview');
    const totalCell = document.getElementById('totalCell');
    const addExpensesMonthlyTotal = document.getElementById('addExpensesMonthlyTotal');
    const expenseDetailModal = document.getElementById('expenseDetailModal');
    const expenseDetailBody = document.getElementById('expenseDetailBody');
    const closeExpenseDetailBtn = document.getElementById('closeExpenseDetailBtn');
    const editExpenseDetailBtn = document.getElementById('editExpenseDetailBtn');
    const saveExpenseDetailBtn = document.getElementById('saveExpenseDetailBtn');
    const cancelExpenseEditBtn = document.getElementById('cancelExpenseEditBtn');
    const deleteExpenseDetailBtn = document.getElementById('deleteExpenseDetailBtn');

    const totalsPeople = document.getElementById('totalsPeople');
    const totalsCategories = document.getElementById('totalsCategories');
    const settlements = document.getElementById('settlements');

    const clearMonthBtn = document.getElementById('clearMonthBtn');
    const checkDuplicatesBtn = document.getElementById('checkDuplicatesBtn');
    const loadBackupsBtn = document.getElementById('loadBackupsBtn');
    const migrateDatesBtn = document.getElementById('migrateDatesBtn');
    const backupStatus = document.getElementById('backupStatus');
    const backupList = document.getElementById('backupList');

    // Spreadsheet import
    const importFile = document.getElementById('importFile');
    const importMapping = document.getElementById('importMapping');
    const importDateColumn = document.getElementById('importDateColumn');
    const importDescriptionColumn = document.getElementById('importDescriptionColumn');
    const importAmountColumn = document.getElementById('importAmountColumn');
    const importPayerColumn = document.getElementById('importPayerColumn');
    const importFormatNotice = document.getElementById('importFormatNotice');
    const importUseSamePayer = document.getElementById('importUseSamePayer');
    const importSamePayer = document.getElementById('importSamePayer');
    const buildImportPreviewBtn = document.getElementById('buildImportPreviewBtn');
    const clearImportBtn = document.getElementById('clearImportBtn');
    const importPreview = document.getElementById('importPreview');
    const categorySuggestions = document.getElementById('categorySuggestions');

    // Analysis
    const analysisView = document.getElementById('analysisView');
    const analysisCompareA = document.getElementById('analysisCompareA');
    const analysisCompareB = document.getElementById('analysisCompareB');
    const analysisSummary = document.getElementById('analysisSummary');
    const analysisTable = document.getElementById('analysisTable');

    // Auth selectors
    const authEmail = document.getElementById('authEmail');
    const authPassword = document.getElementById('authPassword');
    const signInBtn = document.getElementById('signInBtn');
    const signUpBtn = document.getElementById('signUpBtn');
    const signOutBtn = document.getElementById('signOutBtn');
    const forgotPasswordBtn = document.getElementById('forgotPasswordBtn');
    const authStatus = document.getElementById('authStatus');
    const authMessage = document.getElementById('authMessage');
    const themeToggle = document.getElementById('themeToggle');
    const accountCardTitle = document.getElementById('accountCardTitle');
    const accountSectionTitle = document.getElementById('accountSectionTitle');

    // ---------- App state ----------
    let data = null; // current month data (only expenses and lastEdited)
    let fileData = { people: [], categories: [], months: {} }; // Global people/categories + months
    let saveDebounce = null;
    let lastLoadedMonth = null;
    let activeMonth = null;
    let activeExpensePayerFilter = 'all';
    let duplicateReviewState = null;
    let importRowsRaw = [];
    let importDraftRows = [];
    let importDetectedFormat = 'generic';
    let importInProgress = false;
    let duplicateDeleteInProgress = false;
    let cloudSaveQueue = Promise.resolve();
    let cloudSavePending = false;
    let cloudSaveRunning = false;
    let remoteLoadState = 'signed-out';
    let lastCloudSaveAt = null;
    let lastCloudError = '';
    let lastCloudVerifyMessage = '';
    let activeView = 'account';
    let supabaseClient = null;
    let currentUser = null;
    let currentSession = null;

    // ---------- Supabase config ----------
    //
    // SETUP INSTRUCTIONS:
    // 1. Go to https://supabase.com and create a free account
    // 2. Create a new project (choose a region close to you)
    // 3. Once created, go to Project Settings > API
    // 4. Copy the "Project URL" and replace SUPABASE_URL below
    // 5. Copy the "anon public" key and replace SUPABASE_ANON_KEY below
    // 6. Go to SQL Editor and run this SQL to create the required table:
    //
    //    CREATE TABLE family_account_data (
    //      user_id UUID PRIMARY KEY REFERENCES auth.users(id) ON DELETE CASCADE,
    //      data JSONB NOT NULL DEFAULT '{"months": {}}'::jsonb,
    //      updated_at TIMESTAMPTZ DEFAULT NOW()
    //    );
    //
    //    -- Enable Row Level Security
    //    ALTER TABLE family_account_data ENABLE ROW LEVEL SECURITY;
    //
    //    -- Policy: Users can only access their own data
    //    CREATE POLICY "Users can manage their own data"
    //      ON family_account_data
    //      FOR ALL
    //      USING (auth.uid() = user_id)
    //      WITH CHECK (auth.uid() = user_id);
    //
    //    CREATE TABLE family_account_months (
    //      user_id UUID NOT NULL REFERENCES auth.users(id) ON DELETE CASCADE,
    //      month TEXT NOT NULL,
    //      data JSONB NOT NULL DEFAULT '{"expenses": []}'::jsonb,
    //      updated_at TIMESTAMPTZ DEFAULT NOW(),
    //      PRIMARY KEY (user_id, month)
    //    );
    //
    //    ALTER TABLE family_account_months ENABLE ROW LEVEL SECURITY;
    //
    //    CREATE POLICY "Users can manage their own month data"
    //      ON family_account_months
    //      FOR ALL
    //      USING (auth.uid() = user_id)
    //      WITH CHECK (auth.uid() = user_id);
    //
    // 7. Go to Authentication > Settings and configure:
    //    - Enable "Email confirmations" (recommended)
    //    - Set "Site URL" to your app's URL
    //    - Configure email templates if desired
    //
    const SUPABASE_URL = 'https://huxaxrxoacmvtdmxkpsr.supabase.co';
    const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imh1eGF4cnhvYWNtdnRkbXhrcHNyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjY0MTc0OTEsImV4cCI6MjA4MTk5MzQ5MX0.HsG-Q20d3uT5kldXDeJFt6aJmXfxX_S7XWCj9iovRnw';

    // ---------- Utilities ----------
    function thisMonthValue(){
      const d = new Date();
      const m = String(d.getMonth()+1).padStart(2,'0');
      return `${d.getFullYear()}-${m}`;
    }
    function keyFor(month){ return `famacct:${month}` }
    function defaultData(){
      return {expenses:[], lastEdited: new Date().toISOString()};
    }
    function normalizeMonthKey(month){
      const raw = String(month || '').trim();
      const match = /^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?$/.exec(raw);
      if(match){
        return `${match[1]}-${match[2].padStart(2,'0')}`;
      }
      return raw;
    }
    function normalizeMonthData(raw){
      let payload = raw;
      if(typeof payload === 'string'){
        try{
          payload = JSON.parse(payload);
        }catch{
          payload = {};
        }
      }
      if(payload && typeof payload === 'object' && payload.data && typeof payload.data === 'object'){
        if(Array.isArray(payload.data.expenses)){
          payload = payload.data;
        }
      }
      const safe = (payload && typeof payload === 'object') ? { ...payload } : {};
      safe.expenses = Array.isArray(safe.expenses) ? safe.expenses : [];
      if(!safe.lastEdited){ safe.lastEdited = new Date().toISOString(); }
      return safe;
    }
    function hasExpenses(monthData){
      return Array.isArray(monthData?.expenses) && monthData.expenses.length > 0;
    }
    function expenseCount(monthData){
      return Array.isArray(monthData?.expenses) ? monthData.expenses.length : 0;
    }
    function shouldReplaceMonth(existing, incoming){
      if(!existing) return true;
      if(!incoming) return false;
      return !hasExpenses(existing) && hasExpenses(incoming);
    }
    function isNewerMonthData(localData, remoteData){
      if(!localData) return false;
      if(!remoteData) return hasExpenses(localData);
      if(hasExpenses(remoteData) && !hasExpenses(localData)) return false;
      const localTime = Date.parse(localData.lastEdited || '') || 0;
      const remoteTime = Date.parse(remoteData.lastEdited || '') || 0;
      if(localTime && remoteTime && localTime > remoteTime) return true;
      const localCount = expenseCount(localData);
      const remoteCount = expenseCount(remoteData);
      return localCount > remoteCount;
    }
    function mergeProtectedMonths(remoteMonths, localMonths, options = {}){
      const merged = {};
      const allKeys = new Set([
        ...Object.keys(remoteMonths || {}),
        ...Object.keys(localMonths || {})
      ].map(normalizeMonthKey));
      for(const key of allKeys){
        const remoteData = remoteMonths?.[key] ? normalizeMonthData(remoteMonths[key]) : null;
        const localData = localMonths?.[key] ? normalizeMonthData(localMonths[key]) : null;
        if(remoteData && hasExpenses(remoteData) && localData && !hasExpenses(localData) && !options.allowEmptyOverwrite){
          merged[key] = remoteData;
          continue;
        }
        merged[key] = localData || remoteData || defaultData();
      }
      return merged;
    }
    function storeCloudSafetyBackup(remoteData){
      try{
        const backups = JSON.parse(localStorage.getItem('famacct:cloudBackups') || '[]');
        backups.unshift({
          savedAt: new Date().toISOString(),
          data: remoteData || {}
        });
        localStorage.setItem('famacct:cloudBackups', JSON.stringify(backups.slice(0, 20)));
      }catch(err){
        console.warn('Could not create local cloud safety backup', err);
      }
    }
    function getLocalMonthData(monthKey){
      const raw = localStorage.getItem(keyFor(normalizeMonthKey(monthKey)));
      if(!raw) return null;
      try{
        const parsed = JSON.parse(raw);
        return normalizeMonthData(parsed);
      }catch{
        return null;
      }
    }

    function findMonthKeyInMap(monthKey){
      const normalized = normalizeMonthKey(monthKey);
      const months = fileData?.months || {};
      if(months[normalized]) return normalized;
      const keys = Object.keys(months);
      const match = keys.find(k => normalizeMonthKey(k) === normalized);
      return match || normalized;
    }
    async function ensureFreshSession(){
      if(!supabaseClient) return false;
      if(currentSession?.refresh_token){
        try{
          const { data, error } = await supabaseClient.auth.refreshSession({ refresh_token: currentSession.refresh_token });
          if(!error && data?.session){
            currentSession = data.session;
            currentUser = data.session.user;
            return true;
          }
        }catch(err){
          console.warn('Failed to refresh session', err);
        }
      }
      await refreshSession();
      return !!currentSession?.access_token;
    }

    async function getAuthHeaders(){
      if(!currentSession?.access_token && supabaseClient){
        const { data: sessionData } = await supabaseClient.auth.getSession();
        if(sessionData?.session){
          currentSession = sessionData.session;
          currentUser = sessionData.session.user;
        }
      }
      if(!currentSession?.access_token) return null;
      return {
        apikey: SUPABASE_ANON_KEY,
        Authorization: `Bearer ${currentSession.access_token}`,
        Accept: 'application/json'
      };
    }
    async function fetchJson(url, options, retryAuth = true){
      const res = await withTimeout(fetch(url, options), 8000, 'Network request');
      if(res.status === 401 && retryAuth){
        const refreshed = await ensureFreshSession();
        if(refreshed){
          return await fetchJson(url, options, false);
        }
      }
      if(!res.ok){
        const text = await res.text();
        throw new Error(`HTTP ${res.status}: ${text}`);
      }
      const text = await res.text();
      if(!text){
        return null;
      }
      try{
        return JSON.parse(text);
      }catch(err){
        throw new Error(`Invalid JSON response: ${text}`);
      }
    }
    async function fetchMonthsRows(){
      const headers = await getAuthHeaders();
      if(!headers || !currentUser) return [];
      const url = `${SUPABASE_URL}/rest/v1/family_account_months?select=month,data&user_id=eq.${encodeURIComponent(currentUser.id)}`;
      try{
        return await fetchJson(url, { headers });
      }catch(err){
        console.warn('Failed to fetch month rows', err);
        return [];
      }
    }
    async function fetchMonthRow(month, throwOnError = false){
      const headers = await getAuthHeaders();
      if(!headers || !currentUser) return null;
      const key = normalizeMonthKey(month);
      const url = `${SUPABASE_URL}/rest/v1/family_account_months?select=data&user_id=eq.${encodeURIComponent(currentUser.id)}&month=eq.${encodeURIComponent(key)}&limit=1`;
      try{
        const rows = await fetchJson(url, { headers });
        if(Array.isArray(rows) && rows.length && rows[0]?.data){
          return rows[0].data;
        }
      }catch(err){
        if(throwOnError) throw err;
        console.warn('Failed to fetch month row', err);
      }
      return null;
    }
    async function upsertMonthRow(month, monthData){
      const headers = await getAuthHeaders();
      if(!headers || !currentUser) throw new Error('Not signed in.');
      const key = normalizeMonthKey(month);
      const url = `${SUPABASE_URL}/rest/v1/family_account_months?on_conflict=user_id,month`;
      const body = JSON.stringify([{
        user_id: currentUser.id,
        month: key,
        data: monthData,
        updated_at: new Date().toISOString()
      }]);
      await fetchJson(url, {
        method: 'POST',
        headers: {
          ...headers,
          'Content-Type': 'application/json',
          Prefer: 'resolution=merge-duplicates,return=minimal'
        },
        body
      });
    }
    function euro(n){ return n.toLocaleString('es-ES',{style:'currency',currency:'EUR'}) }
    function round2(n){ return Math.round((n+Number.EPSILON)*100)/100 }
    function debounceSave(){ clearTimeout(saveDebounce); saveDebounce = setTimeout(()=>persist(), 400); }
    function withTimeout(promise, ms, label){
      let timeoutId;
      const timeout = new Promise((_, reject)=>{
        timeoutId = setTimeout(()=>{
          const message = label ? `${label} timed out` : 'Operation timed out';
          reject(new Error(message));
        }, ms);
      });
      return Promise.race([promise, timeout]).finally(()=> clearTimeout(timeoutId));
    }
    function setModeBadge(){
      if(currentUser){
        modeBadge.textContent = `Signed in: ${currentUser.email}`;
        return;
      }
      modeBadge.textContent = 'Local storage';
    }
    function setCloudStatus(status, message){
      if(cloudSyncStatus) cloudSyncStatus.textContent = message;
      if(saveCloudBtn){
        saveCloudBtn.disabled = ['signed-out', 'syncing', 'loading', 'failed'].includes(status) || !currentUser;
        saveCloudBtn.textContent = status === 'syncing' ? 'Saving...' : 'Save to cloud';
      }
      if(retryCloudBtn){
        retryCloudBtn.disabled = !currentUser || status === 'syncing' || status === 'loading';
      }
      if(cloudStatusDetail){
        let detail = '';
        if(!currentUser){
          detail = 'Local-only mode. Changes stay on this device until you sign in.';
        } else if(status === 'loading'){
          detail = 'Loading cloud data. Uploads are blocked until this finishes.';
        } else if(status === 'failed'){
          detail = `Cloud data did not load. Uploads are blocked to protect existing data.${lastCloudError ? ' Last error: ' + lastCloudError : ''}`;
        } else if(status === 'syncing'){
          detail = 'Saving local changes to Supabase now.';
        } else if(status === 'pending'){
          detail = `Saved locally. Cloud retry is pending.${lastCloudError ? ' Last error: ' + lastCloudError : ''}`;
        } else {
          const savedAt = lastCloudSaveAt ? formatUK(lastCloudSaveAt) : 'not yet in this session';
          detail = `Cloud ready. Last confirmed cloud save: ${savedAt}.${lastCloudVerifyMessage ? ' ' + lastCloudVerifyMessage : ''}`;
        }
        cloudStatusDetail.textContent = detail;
      }
      const blockEditing = !!currentUser && (status === 'loading' || status === 'failed');
      [
        addQuickBtn,
        saveBatchBtn,
        buildImportPreviewBtn,
        clearMonthBtn,
        checkDuplicatesBtn,
        migrateDatesBtn
      ].forEach(button => {
        if(button) button.disabled = blockEditing;
      });
    }
    function updateCloudStatus(){
      if(!currentUser){
        remoteLoadState = 'signed-out';
        setCloudStatus('signed-out', 'Cloud: not signed in');
      } else if(remoteLoadState === 'loading'){
        setCloudStatus('loading', 'Cloud: loading...');
      } else if(remoteLoadState === 'failed'){
        setCloudStatus('failed', 'Cloud: load failed');
      } else if(cloudSaveRunning){
        setCloudStatus('syncing', 'Cloud: syncing...');
      } else if(cloudSavePending){
        setCloudStatus('pending', 'Cloud: local changes pending');
      } else {
        setCloudStatus('synced', 'Cloud: synced');
      }
    }
    function queueCloudSave(label, task){
      if(!currentUser){ updateCloudStatus(); return; }
      if(remoteLoadState !== 'ready'){
        cloudSavePending = true;
        updateCloudStatus();
        console.warn(`${label}: cloud data is not loaded, so upload was blocked.`);
        return;
      }
      cloudSavePending = true;
      updateCloudStatus();
      cloudSaveQueue = cloudSaveQueue
        .catch(()=>{})
        .then(async ()=>{
          cloudSaveRunning = true;
          updateCloudStatus();
          try{
            await withTimeout(task(), 9000, label);
            cloudSavePending = false;
            lastCloudError = '';
            lastCloudSaveAt = new Date().toISOString();
          }catch(err){
            console.warn(label, err);
            lastCloudError = err.message || String(err);
            cloudSavePending = true;
          }finally{
            cloudSaveRunning = false;
            updateCloudStatus();
          }
        });
    }
    async function syncCloudNow(options = {}){
      if(!currentUser){
        alert('Sign in before saving to cloud.');
        updateCloudStatus();
        return false;
      }
      if(remoteLoadState !== 'ready'){
        alert('Cloud data has not loaded successfully yet. To protect your existing data, nothing will be uploaded until cloud loading succeeds.');
        updateCloudStatus();
        return false;
      }
      const normalizedMonth = normalizeMonthKey(activeMonth || monthEl.value);
      const snapshot = JSON.parse(JSON.stringify(data || defaultData()));
      localSaveCurrent(normalizedMonth, snapshot);
      data = snapshot;
      fileData.months[normalizedMonth] = snapshot;
      cloudSaveRunning = true;
      updateCloudStatus();
      try{
        await withTimeout((async ()=>{
          await cloudSaveQueue.catch(()=>{});
          await saveRemote(options);
          await saveRemoteMonth(normalizedMonth, snapshot, options);
          const savedMonth = await loadRemoteMonthStrict(normalizedMonth);
          const savedCount = (savedMonth?.expenses || []).length;
          const expectedCount = (snapshot?.expenses || []).length;
          if(savedCount !== expectedCount){
            throw new Error(`Cloud verification failed for ${normalizedMonth}: expected ${expectedCount} expenses, found ${savedCount}.`);
          }
          lastCloudVerifyMessage = `${normalizedMonth} verified with ${savedCount} expenses.`;
        })(), 12000, 'Manual cloud save');
        cloudSavePending = false;
        lastCloudError = '';
        lastCloudSaveAt = new Date().toISOString();
        updateCloudStatus();
        return true;
      }catch(err){
        console.warn('Manual cloud save failed', err);
        cloudSavePending = true;
        updateCloudStatus();
        alert(`Cloud save failed. Your changes are still saved on this device.\n\n${err.message || err}`);
        return false;
      }finally{
        cloudSaveRunning = false;
        updateCloudStatus();
      }
    }

    async function retryCloudSync(){
      if(!currentUser){
        alert('Sign in before retrying cloud sync.');
        return;
      }
      if(remoteLoadState !== 'ready'){
        const loaded = await loadRemote();
        if(!loaded){
          alert('Cloud data still could not be loaded. Uploads remain blocked.');
          return;
        }
        await loadFromFileOrLocal(monthEl.value);
        renderAll();
      }
      await syncCloudNow();
    }
    function formatUK(iso){
      if(!iso) return '—';
      const d = new Date(iso);
      const dd = String(d.getDate()).padStart(2,'0');
      const mm = String(d.getMonth()+1).padStart(2,'0');
      const yyyy = d.getFullYear();
      const HH = String(d.getHours()).padStart(2,'0');
      const min = String(d.getMinutes()).padStart(2,'0');
      return `${dd}/${mm}/${yyyy} ${HH}:${min}`;
    }
    function todayDateValue(){
      return new Date().toISOString().slice(0, 10);
    }
    function dateValueFromMonth(month){
      const key = normalizeMonthKey(month);
      return /^\d{4}-\d{2}$/.test(key) ? `${key}-01` : todayDateValue();
    }
    function formatExpenseDate(value){
      if(!value) return '—';
      const parsed = parseImportDate(value, /^\d{1,2}\/\d{1,2}\/\d{2,4}$/.test(String(value || '')));
      if(!parsed) return String(value);
      return parsed.toLocaleDateString('en-GB');
    }
    function dateInputValueFromImport(value, preferMonthFirst = false){
      const parsed = parseImportDate(value, preferMonthFirst);
      if(!parsed) return '';
      return `${parsed.getFullYear()}-${String(parsed.getMonth() + 1).padStart(2, '0')}-${String(parsed.getDate()).padStart(2, '0')}`;
    }
    function setLastEditedNow(){ if(data){ data.lastEdited = new Date().toISOString(); renderLastEdited(); } }
    function renderLastEdited(){ lastEditedEl.textContent = formatUK(data?.lastEdited); }

    function applyTheme(mode){
      const isDark = mode === 'dark';
      document.body.classList.toggle('dark', isDark);
      if(themeToggle){ themeToggle.checked = isDark; }
    }
    function loadTheme(){
      const stored = localStorage.getItem('famacct:theme');
      applyTheme(stored === 'dark' ? 'dark' : 'light');
    }
    function saveTheme(isDark){
      localStorage.setItem('famacct:theme', isDark ? 'dark' : 'light');
      applyTheme(isDark ? 'dark' : 'light');
    }

    // ---------- Auth utilities ----------
    function showAuthMessage(message, type = 'info'){
      // type: 'success', 'error', 'info'
      authMessage.textContent = message;
      authMessage.style.display = 'block';
      if(type === 'success'){
        authMessage.style.background = '#ecfdf5';
        authMessage.style.color = '#065f46';
        authMessage.style.border = '1px solid #6ee7b7';
      } else if(type === 'error'){
        authMessage.style.background = '#fef2f2';
        authMessage.style.color = '#7f1d1d';
        authMessage.style.border = '1px solid #fca5a5';
      } else {
        authMessage.style.background = '#eff6ff';
        authMessage.style.color = '#1e3a8a';
        authMessage.style.border = '1px solid #bfdbfe';
      }
      // Auto-hide after 8 seconds for success/info messages
      if(type !== 'error'){
        setTimeout(() => { authMessage.style.display = 'none'; }, 8000);
      }
    }
    function hideAuthMessage(){ authMessage.style.display = 'none'; }
    function validateEmail(email){
      const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      return re.test(email);
    }
    function validatePassword(password){
      return password && password.length >= 6;
    }
    function setButtonLoading(button, isLoading){
      if(isLoading){
        button.disabled = true;
        button.dataset.originalText = button.textContent;
        button.textContent = '⏳ Loading...';
      } else {
        button.disabled = false;
        if(button.dataset.originalText){
          button.textContent = button.dataset.originalText;
          delete button.dataset.originalText;
        }
      }
    }

    // ---------- Collapsible cards ----------
    const COLLAPSE_KEY = 'famacct:collapsed';
    function loadCollapsed(){ try{ return JSON.parse(localStorage.getItem(COLLAPSE_KEY)||'{}'); }catch{return {}}; }
    function saveCollapsed(obj){ try{ localStorage.setItem(COLLAPSE_KEY, JSON.stringify(obj)); }catch{} }
    function setupCollapsibles(){
      const state = loadCollapsed();
      document.querySelectorAll('.card[data-section]').forEach(card=>{
        const key = card.getAttribute('data-section');
        const header = card.querySelector('.card-header');
        const content = card.querySelector('.card-content');
        const collapsed = !!state[key];
        if(collapsed){ card.classList.add('collapsed'); content.style.display = 'none'; }
        header.addEventListener('click', (e)=>{
          if(e.target.closest('button')) return; // don’t collapse when clicking a button inside header
          const now = card.classList.toggle('collapsed');
          content.style.display = now ? 'none' : 'block';
          const st = loadCollapsed(); st[key] = now ? 1 : 0; saveCollapsed(st);
        });
      });
    }

    // ---------- Storage: Local ----------
    function localLoadMonth(month){
      const raw = localStorage.getItem(keyFor(month));
      if(raw){
        try{
          const parsed = JSON.parse(raw);
          // Migrate old format: extract people/categories if they exist
          if(parsed.people || parsed.categories){
            // Merge into global lists
            if(parsed.people){
              fileData.people = Array.from(new Set([...(fileData.people||[]), ...(parsed.people||[])]));
            }
            if(parsed.categories){
              fileData.categories = Array.from(new Set([...(fileData.categories||[]), ...(parsed.categories||[])]));
            }
            // Save migrated data
            saveGlobalToLocal();
          }
          // Only keep expenses and lastEdited for month data
          data = {
            expenses: parsed.expenses || [],
            lastEdited: parsed.lastEdited || new Date().toISOString()
          };
        }catch{
          data = defaultData();
        }
        if(!data.lastEdited){ data.lastEdited = new Date().toISOString(); }
      } else {
        data = defaultData();
        localStorage.setItem(keyFor(month), JSON.stringify(data));
      }
      activeMonth = normalizeMonthKey(month);
    }
    function localSaveCurrent(monthKey, payload){
      // Only save expenses and lastEdited for each month
      const key = keyFor(monthKey || monthEl.value);
      localStorage.setItem(key, JSON.stringify(payload || data));
    }
    function saveGlobalToLocal(){
      // Save global people and categories
      try{
        localStorage.setItem('famacct:global', JSON.stringify({
          people: fileData.people || [],
          categories: fileData.categories || []
        }));
      }catch(e){
        console.error('Failed to save global data', e);
      }
    }
    function loadGlobalFromLocal(){
      // Load global people and categories
      try{
        const raw = localStorage.getItem('famacct:global');
        if(raw){
          const global = JSON.parse(raw);
          fileData.people = global.people || [];
          fileData.categories = global.categories || [];
        } else {
          // No global data exists yet - migrate from all saved months
          console.log('Migrating people and categories from all saved months...');
          const allPeople = new Set();
          const allCategories = new Set();

          // Scan all saved months and extract people/categories
          for(let i=0; i<localStorage.length; i++){
            const key = localStorage.key(i);
            if(key && key.startsWith('famacct:') && key !== 'famacct:global' && key !== 'famacct:collapsed'){
              try{
                const monthData = JSON.parse(localStorage.getItem(key));
                if(monthData.people && Array.isArray(monthData.people)){
                  monthData.people.forEach(p => allPeople.add(String(p).trim()));
                }
                if(monthData.categories && Array.isArray(monthData.categories)){
                  monthData.categories.forEach(c => allCategories.add(String(c).trim()));
                }
              }catch(e){
                console.error('Failed to parse month data for migration:', key, e);
              }
            }
          }

          fileData.people = Array.from(allPeople).filter(Boolean);
          fileData.categories = Array.from(allCategories).filter(Boolean);

          // Save the migrated global data
          if(fileData.people.length > 0 || fileData.categories.length > 0){
            console.log('Migrated', fileData.people.length, 'people and', fileData.categories.length, 'categories');
            saveGlobalToLocal();
          }
        }
      }catch(e){
        console.error('Failed to load global data', e);
      }
    }
    function listSavedMonthsLocal(){
      const months = [];
      for(let i=0;i<localStorage.length;i++){
        const k = localStorage.key(i);
        if(k && /^famacct:\d{4}-\d{2}$/.test(k)) months.push(k.split(':')[1]);
      }
      months.sort((a,b)=> a<b?1:(a>b?-1:0));
      return months;
    }

    function listSavedMonthsRemote(){
      const months = Object.keys(fileData.months||{});
      const seen = new Set();
      const normalized = [];
      for(const m of months){
        const key = normalizeMonthKey(m);
        if(seen.has(key)) continue;
        seen.add(key);
        normalized.push(key);
      }
      normalized.sort((a,b)=> a<b?1:(a>b?-1:0));
      return normalized;
    }

    // ---------- Shared load/save ----------
    function monthExists(month){
      const normalizedMonth = normalizeMonthKey(month);
      if(currentUser){
        if(fileData.months){
          const mapKey = findMonthKeyInMap(normalizedMonth);
          if(fileData.months[mapKey]) return true;
        }
      }
      return localStorage.getItem(keyFor(normalizedMonth)) != null;
    }

    async function loadFromFileOrLocal(month){
      const normalizedMonth = normalizeMonthKey(month);
      activeMonth = normalizedMonth;
      if(currentUser){
        const mapKey = findMonthKeyInMap(normalizedMonth);
        const existing = (fileData.months && fileData.months[mapKey]) || null;
        if(existing){
          const normalizedExisting = normalizeMonthData(existing);
          if(!hasExpenses(normalizedExisting)){
            const remoteData = await loadRemoteMonth(normalizedMonth);
            const normalizedRemote = remoteData ? normalizeMonthData(remoteData) : null;
            if(shouldReplaceMonth(normalizedExisting, normalizedRemote)){
              data = normalizedRemote;
              fileData.months[normalizedMonth] = data;
              return;
            }
          }
          data = normalizedExisting;
          fileData.months[normalizedMonth] = data;
          return;
        }
        const remoteData = await loadRemoteMonth(normalizedMonth);
        if(remoteData){
          data = normalizeMonthData(remoteData);
          fileData.months[normalizedMonth] = data;
          return;
        }
        const localRaw = localStorage.getItem(keyFor(normalizedMonth));
        if(localRaw){
          try{
            const parsed = JSON.parse(localRaw);
            // Only keep expenses and lastEdited
            data = {
              expenses: parsed.expenses || [],
              lastEdited: parsed.lastEdited || new Date().toISOString()
            };
          }catch{
            data = defaultData();
          }
        } else {
          data = defaultData();
        }
      } else {
        localLoadMonth(normalizedMonth);
      }
    }

    async function persist(options = {}){
      const targetMonth = activeMonth || monthEl.value;
      const normalizedMonth = normalizeMonthKey(targetMonth);
      const snapshot = JSON.parse(JSON.stringify(data || defaultData()));
      try{ localSaveCurrent(normalizedMonth, snapshot); }catch{}
      data = snapshot;
      if(currentUser){
        fileData.months[normalizedMonth] = snapshot;
        queueCloudSave('Cloud sync failed', async ()=>{
          await saveRemote(options);
          await saveRemoteMonth(normalizedMonth, snapshot, options);
        });
      } else {
        saveGlobalToLocal();
      }
      renderAverages();
    }

    // ---------- Helpers ----------
    function uniqueList(arr){ return Array.from(new Set((arr||[]).map(s=>String(s).trim()).filter(Boolean))); }

    // ---------- Split helpers ----------
    function buildSplitSliders(containerEl){
      const ppl = fileData.people || [];
      containerEl.innerHTML='';
      if(ppl.length === 0) return;

      if(ppl.length === 1){
        const row = document.createElement('div');
        row.className = 'split-single';
        row.dataset.person = ppl[0];
        row.textContent = `${ppl[0]} 100%`;
        containerEl.appendChild(row);
        return;
      }

      if(ppl.length === 2){
        const [leftPerson, rightPerson] = ppl;
        const row = document.createElement('div');
        row.className = 'two-person-split';

        const leftLabel = document.createElement('span');
        leftLabel.className = 'split-end split-left';

        const slider = document.createElement('input');
        slider.type = 'range';
        slider.min = '0';
        slider.max = '100';
        slider.step = '5';
        slider.value = '50';
        slider.className = 'split-slider two-person-split-slider';
        slider.dataset.leftPerson = leftPerson;
        slider.dataset.rightPerson = rightPerson;

        const rightLabel = document.createElement('span');
        rightLabel.className = 'split-end split-right';

        const updateTwoPersonLabels = () => {
          const leftValue = parseInt(slider.value) || 0;
          leftLabel.textContent = `${leftPerson} ${leftValue}%`;
          rightLabel.textContent = `${100 - leftValue}% ${rightPerson}`;
        };

        slider.addEventListener('input', updateTwoPersonLabels);
        updateTwoPersonLabels();

        row.appendChild(leftLabel);
        row.appendChild(slider);
        row.appendChild(rightLabel);
        containerEl.appendChild(row);
        return;
      }

      // Create a slider for each person
      ppl.forEach((person, index) => {
        const row = document.createElement('div');
        row.className = 'split-slider-row';

        const label = document.createElement('label');
        label.textContent = person;

        const slider = document.createElement('input');
        slider.type = 'range';
        slider.min = '0';
        slider.max = '100';
        slider.step = '5';
        slider.value = Math.floor(100 / ppl.length);
        slider.dataset.person = person;
        slider.className = 'split-slider';

        const valueDisplay = document.createElement('span');
        valueDisplay.className = 'split-value';
        valueDisplay.textContent = slider.value + '%';

        // Update display when slider changes
        slider.addEventListener('input', (e) => {
          valueDisplay.textContent = e.target.value + '%';
          // Auto-adjust other sliders to maintain 100% total
          autoAdjustSliders(containerEl, person);
        });

        row.appendChild(label);
        row.appendChild(slider);
        row.appendChild(valueDisplay);
        containerEl.appendChild(row);
      });
    }

    function autoAdjustSliders(containerEl, changedPerson){
      const sliders = containerEl.querySelectorAll('.split-slider');
      const values = {};
      let total = 0;

      // Get all current values
      sliders.forEach(slider => {
        const val = parseInt(slider.value);
        values[slider.dataset.person] = val;
        total += val;
      });

      // If total is not 100, adjust other sliders proportionally
      if(total !== 100 && sliders.length > 1){
        const diff = total - 100;
        const others = Array.from(sliders).filter(s => s.dataset.person !== changedPerson);

        if(others.length > 0){
          // Distribute the difference among other sliders
          const adjustment = Math.floor(diff / others.length);
          let remaining = diff;

          others.forEach((slider, index) => {
            const currentVal = parseInt(slider.value);
            const adj = (index === others.length - 1) ? remaining : adjustment;
            const newVal = Math.max(0, Math.min(100, currentVal - adj));
            slider.value = newVal;
            slider.nextElementSibling.textContent = newVal + '%';
            remaining -= adj;
          });
        }
      }
    }

    function getSplitsFromSliders(containerEl){
      const splits = {};
      const twoPersonSlider = containerEl.querySelector('.two-person-split-slider');
      if(twoPersonSlider){
        const leftValue = parseInt(twoPersonSlider.value) || 0;
        splits[twoPersonSlider.dataset.leftPerson] = leftValue;
        splits[twoPersonSlider.dataset.rightPerson] = 100 - leftValue;
        return splits;
      }
      const singlePerson = containerEl.querySelector('.split-single');
      if(singlePerson?.dataset.person){
        splits[singlePerson.dataset.person] = 100;
        return splits;
      }
      const sliders = containerEl.querySelectorAll('.split-slider');
      sliders.forEach(slider => {
        splits[slider.dataset.person] = parseInt(slider.value);
      });
      return splits;
    }
    function splitsToLabel(splits){
      const parts = Object.entries(splits||{}).filter(([_,v])=>v>0).map(([k,v])=>`${k} ${v}%`);
      return parts.join(' / ') || '—';
    }

    // ---------- Computations & Rendering ----------
    // CHANGED: computeTotals now shows what each person ACTUALLY PAID (sum by payer), not split shares
    function computeTotals(){
      const perPerson = Object.fromEntries((fileData.people||[]).map(p=>[p,0]));
      const sharedPaidByPerson = Object.fromEntries((fileData.people||[]).map(p=>[p,0]));
      const perCategory = Object.fromEntries((fileData.categories||[]).map(c=>[c,0]));
      let grand=0;
      let sharedGrand=0;
      for(const e of (data.expenses||[])){
        const amount = +e.amount || 0;
        const payer = e.payer || e.person || '';
        if(!(payer in perPerson)) perPerson[payer]=0;
        if(!(payer in sharedPaidByPerson)) sharedPaidByPerson[payer]=0;
        perPerson[payer] = round2((perPerson[payer]||0) + amount);
        const splits = e.splits || {[payer]:100};
        const sharedExpense = Object.values(splits).filter(pct => (+pct || 0) > 0).length > 1;
        if(sharedExpense){
          sharedPaidByPerson[payer] = round2((sharedPaidByPerson[payer] || 0) + amount);
          sharedGrand = round2(sharedGrand + amount);
        }
        if(!(e.category in perCategory)) perCategory[e.category]=0;
        perCategory[e.category]+= amount;
        grand+= amount;
      }
      perPerson.__total = round2(grand);
      sharedPaidByPerson.__total = round2(sharedGrand);
      return {perPerson, sharedPaidByPerson, perCategory};
    }

    // CHANGED: computeSettlements uses PAID vs OWED (OWED from splits), not equal-share-of-grand-total
    function computeSettlements(){
      const people = (fileData.people||[]);
      const paid = Object.fromEntries(people.map(p=>[p,0]));
      const owed = Object.fromEntries(people.map(p=>[p,0]));
      let grand = 0;
      let sharedTotal = 0;

      for(const e of (data.expenses||[])){
        const amount = +e.amount || 0;
        const payer = e.payer || e.person || '';
        grand += amount;

        // what was actually paid
        if(!(payer in paid)) paid[payer]=0;
        paid[payer] = round2((paid[payer]||0) + amount);

        // liability owed according to splits
        const splits = e.splits || {[payer]:100};
        const entries = Object.entries(splits);
        const activeSplitPeople = entries.filter(([_, pct]) => (+pct || 0) > 0);
        if(activeSplitPeople.length > 1){
          sharedTotal = round2(sharedTotal + amount);
        }
        let allocated = 0;
        entries.forEach(([p, pct], idx) => {
          if(!(p in owed)) owed[p]=0;
          let part = round2(amount * ((+pct || 0)/100));
          allocated = round2(allocated + part);
          if(idx === entries.length - 1){
            const remainder = round2(amount - allocated);
            part = round2(part + remainder);
          }
          owed[p] = round2((owed[p]||0) + part);
        });
      }

      const balances = people.map(p=>({ person:p, balance: round2((paid[p]||0) - (owed[p]||0)) }));

      // greedy settlement: debtors (negative) pay creditors (positive)
      const creditors = balances.filter(b=>b.balance>0).map(b=>({...b})).sort((a,b)=>b.balance-a.balance);
      const debtors = balances.filter(b=>b.balance<0).map(b=>({...b})).sort((a,b)=>a.balance-b.balance);
      const transfers = [];
      let ci=0, di=0;
      while(ci<creditors.length && di<debtors.length){
        const c = creditors[ci], d = debtors[di];
        const amt = round2(Math.min(c.balance, -d.balance));
        if(amt>0){
          transfers.push({from:d.person,to:c.person,amount:amt});
          c.balance = round2(c.balance - amt);
          d.balance = round2(d.balance + amt);
        }
        if(c.balance<=0.0001) ci++;
        if(d.balance>=-0.0001) di++;
      }
      // Display-only summary: average share of expenses split between more than one person.
      const n = people.length || 1;
      const share = round2((sharedTotal||0)/n);
      return {share, balances, transfers};
    }

    function renderPeopleCats(){
      peopleChips.innerHTML = '';
      for(const p of (fileData.people||[])){
        const c = document.createElement('span');
        c.className='chip';
        c.innerHTML = '<span>'+p+'</span><span class="x" title="Remove">✖</span>';
        c.querySelector('.x').addEventListener('click', async ()=>{ if(confirm('Remove '+p+' and their expenses from ALL months?')) { await removePerson(p); }});
        peopleChips.appendChild(c);
      }
      categoriesChips.innerHTML = '';
      for(const cName of (fileData.categories||[])){
        const c = document.createElement('span');
        c.className='chip';
        const label = document.createElement('span');
        label.textContent = cName;
        const editBtn = document.createElement('button');
        editBtn.className = 'chip-edit';
        editBtn.title = 'Rename category';
        editBtn.type = 'button';
        editBtn.textContent = 'Edit';
        const removeBtn = document.createElement('span');
        removeBtn.className = 'x';
        removeBtn.title = 'Remove';
        removeBtn.textContent = '✖';
        editBtn.addEventListener('click', async ()=>{ await renameCategoryPrompt(cName); });
        removeBtn.addEventListener('click', async ()=>{ if(confirm('Remove category \"'+cName+'\" from ALL months?')) { await removeCategory(cName); }});
        c.appendChild(label);
        c.appendChild(editBtn);
        c.appendChild(removeBtn);
        categoriesChips.appendChild(c);
      }
      fillSelect(quickPerson, (fileData.people||[]));
      fillSelect(quickCategory, (fileData.categories||[]));
      fillSelect(batchPerson, (fileData.people||[]));
      if(importSamePayer) fillSelect(importSamePayer, (fileData.people||[]));
      buildSplitSliders(quickSplitSliders);
      buildSplitSliders(batchSplitSliders);
      [...batchRows.querySelectorAll('.batchCategory')].forEach(sel=>fillSelect(sel, (fileData.categories||[])));
    }
    function renderExpenses(){
      const monthExpenses = data.expenses || [];
      const payerOptions = uniqueList([
        ...(fileData.people || []),
        ...monthExpenses.map(e => e.payer || e.person || '').filter(Boolean)
      ]);
      if(activeExpensePayerFilter !== 'all' && !payerOptions.includes(activeExpensePayerFilter)){
        activeExpensePayerFilter = 'all';
      }
      if(expensePayerTabs){
        expensePayerTabs.innerHTML = '';
        const tabs = [{value:'all', label:'All'}, ...payerOptions.map(person => ({value:person, label:person}))];
        for(const tab of tabs){
          const button = document.createElement('button');
          button.type = 'button';
          button.textContent = tab.label;
          button.className = tab.value === activeExpensePayerFilter ? 'active' : '';
          button.addEventListener('click', () => {
            activeExpensePayerFilter = tab.value;
            renderExpenses();
          });
          expensePayerTabs.appendChild(button);
        }
      }
      const visibleExpenses = activeExpensePayerFilter === 'all'
        ? monthExpenses
        : monthExpenses.filter(e => (e.payer || e.person || '') === activeExpensePayerFilter);
      const headRow = expensesTable.querySelector('thead tr');
      const footRow = expensesTable.querySelector('tfoot tr');

      if(headRow){
        headRow.innerHTML = '';
        ['#', 'Date', 'Payer', 'Category', 'Note', 'Amount', 'Split', ''].forEach(label => {
          const th = document.createElement('th');
          th.textContent = label;
          headRow.appendChild(th);
        });
      }

      if(!monthExpenses.length){
        expensesEmpty.style.display='block';
        expensesTable.style.display='none';
        const currentTotalCell = document.getElementById('totalCell');
        if(currentTotalCell) currentTotalCell.textContent = euro(0);
        if(addExpensesMonthlyTotal) addExpensesMonthlyTotal.textContent = euro(0);
        expensesBody.innerHTML = '';
        return;
      }
      expensesEmpty.style.display = visibleExpenses.length ? 'none' : 'block';
      expensesEmpty.textContent = visibleExpenses.length ? 'No expenses yet.' : 'No expenses for this payer.';
      expensesTable.style.display='table';
      expensesBody.innerHTML = '';
      let i=1, total=0;
      for(const [idx,e] of monthExpenses.entries()){
        const payer = e.payer || e.person || '';
        if(activeExpensePayerFilter !== 'all' && payer !== activeExpensePayerFilter) continue;
        const amount = +e.amount || 0;
        total += amount;
        const tr = document.createElement('tr');
        tr.className = 'expense-row';
        tr.tabIndex = 0;
        tr.addEventListener('click', event => {
          if(event.target.closest('button')) return;
          openExpenseDetail(idx);
        });
        tr.addEventListener('keydown', event => {
          if(event.key === 'Enter' || event.key === ' '){
            event.preventDefault();
            openExpenseDetail(idx);
          }
        });
        const splitTxt = splitsToLabel(e.splits||{[payer]:100});
        const cells = [
          { label:'#', value:i++ },
          { label:'Date', value:`${formatExpenseDate(e.date)}${e.dateEstimated ? ' (estimated)' : ''}` },
          { label:'Payer', value:payer },
          { label:'Category', value:e.category || '' },
          { label:'Note', value:e.note || '' },
          { label:'Amount', value:euro(amount) },
          { label:'Split', value:splitTxt }
        ];
        for(const cell of cells){
          const td = document.createElement('td');
          td.dataset.label = cell.label;
          td.textContent = cell.value;
          tr.appendChild(td);
        }
        const deleteCell = document.createElement('td');
        deleteCell.dataset.label = 'Delete';
        const deleteBtn = document.createElement('button');
        deleteBtn.className = 'btn ghost';
        deleteBtn.title = 'Delete';
        deleteBtn.dataset.i = String(idx);
        deleteBtn.type = 'button';
        deleteBtn.textContent = 'x';
        deleteBtn.addEventListener('click', async ()=>{ (data.expenses||[]).splice(idx,1); await save(); renderAll(); });
        deleteCell.appendChild(deleteBtn);
        tr.appendChild(deleteCell);
        expensesBody.appendChild(tr);
      }
      if(footRow){
        footRow.innerHTML = '';
        const footerCells = [
          { label:'#', value:'' },
          { label:'Date', value:'' },
          { label:'Payer', value:'' },
          { label:'Category', value:'Total' },
          { label:'Note', value:'' },
          { label:'Amount', value:euro(round2(total)), id:'totalCell' },
          { label:'Split', value:'' },
          { label:'Delete', value:'' }
        ];
        for(const cell of footerCells){
          const td = document.createElement('td');
          td.dataset.label = cell.label;
          td.textContent = cell.value;
          if(cell.id) td.id = cell.id;
          footRow.appendChild(td);
        }
      }
      const monthlyTotal = monthExpenses.reduce((sum, expense) => round2(sum + (+expense.amount || 0)), 0);
      if(addExpensesMonthlyTotal) addExpensesMonthlyTotal.textContent = euro(round2(monthlyTotal));
    }
    function renderTotals(){
      const {perPerson, sharedPaidByPerson, perCategory} = computeTotals();
      const total = perPerson.__total || 0;
      const sharedTotal = sharedPaidByPerson.__total || 0;
      let htmlP='';
      if((fileData.people||[]).length){
        htmlP += '<table><thead><tr><th>Person</th><th>Total paid this month</th><th>Total shared costs paid</th></tr></thead><tbody>';
        for(const p of (fileData.people||[])){
          htmlP += '<tr><td>'+p+'</td><td>'+euro(round2(perPerson[p]||0))+'</td><td>'+euro(round2(sharedPaidByPerson[p]||0))+'</td></tr>';
        }
        htmlP += '</tbody><tfoot><tr><td>All people</td><td>'+euro(round2(total))+'</td><td>'+euro(round2(sharedTotal))+'</td></tr></tfoot></table>';
      } else htmlP = '<div class="muted">Add some people to see totals.</div>';
      totalsPeople.innerHTML = htmlP;

      let htmlC='';
      if((fileData.categories||[]).length){
        htmlC += '<table><thead><tr><th>Category</th><th>Total</th></tr></thead><tbody>';
        for(const c of (fileData.categories||[])){
          htmlC += '<tr><td>'+c+'</td><td>'+euro(round2(perCategory[c]||0))+'</td></tr>';
        }
        htmlC += '</tbody><tfoot><tr><td>All categories</td><td>'+euro(round2(total))+'</td></tr></tfoot></table>';
      } else htmlC = '<div class="muted">Add some categories to see totals.</div>';
      totalsCategories.innerHTML = htmlC;
    }
    function renderSettlements(){
      const {share, balances, transfers} = computeSettlements();
      if(!(fileData.people||[]).length){ settlements.innerHTML = '<div class="muted">Add people first.</div>'; return; }
      let html = '<div class="row"><div>Equal share per person (based on splits): <strong>'+euro(share)+'</strong></div></div>';
      html += '<table><thead><tr><th>Person</th><th>Balance (positive = others owe them, negative = they owe others)</th></tr></thead><tbody>';
      for(const b of balances){
        const cls = b.balance>=0? 'ok' : 'bad';
        html += '<tr><td>'+b.person+'</td><td><span class="pill '+cls+'">'+(b.balance>=0?'+':'')+euro(b.balance)+'</span></td></tr>';
      }
      html += '</tbody></table>';
      if(!transfers.length){ html += '<div class="muted">No transfers needed ✅</div>'; }
      else{
        html += '<div class="card-title" style="margin:8px 0">Suggested transfers</div>';
        html += '<table><thead><tr><th>From</th><th>To</th><th>Amount</th></tr></thead><tbody>';
        for(const t of transfers){ html += '<tr><td>'+t.from+'</td><td>'+t.to+'</td><td>'+euro(t.amount)+'</td></tr>'; }
        html += '</tbody></table>';
      }
      settlements.innerHTML = html;
    }
    function escapeHtml(value){
      return String(value || '').replace(/[<>&"]/g, char => ({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;'}[char]));
    }
    function setExpenseDetailMode(mode){
      const editing = mode === 'edit';
      if(editExpenseDetailBtn) editExpenseDetailBtn.style.display = editing ? 'none' : '';
      if(saveExpenseDetailBtn) saveExpenseDetailBtn.style.display = editing ? '' : 'none';
      if(cancelExpenseEditBtn) cancelExpenseEditBtn.style.display = editing ? '' : 'none';
      if(deleteExpenseDetailBtn) deleteExpenseDetailBtn.style.display = editing ? 'none' : '';
    }
    function expenseDetailIndex(){
      const index = Number(expenseDetailModal?.dataset.index);
      return Number.isInteger(index) ? index : -1;
    }
    function openExpenseDetail(index){
      const expense = (data.expenses || [])[index];
      if(!expense || !expenseDetailModal || !expenseDetailBody) return;
      expenseDetailModal.dataset.index = String(index);
      renderExpenseDetailView(index);
      expenseDetailModal.style.display = 'flex';
    }
    function renderExpenseDetailView(index){
      const expense = (data.expenses || [])[index];
      if(!expense || !expenseDetailBody) return;
      const payer = expense.payer || expense.person || '';
      const rows = [
        ['Date', `${formatExpenseDate(expense.date)}${expense.dateEstimated ? ' (estimated)' : ''}`],
        ['Payer', payer],
        ['Category', expense.category || ''],
        ['Note', expense.note || ''],
        ['Amount', euro(+expense.amount || 0)],
        ['Split', splitsToLabel(expense.splits || {[payer]:100})]
      ];
      expenseDetailBody.innerHTML = '<table><tbody>' + rows.map(([label, value]) => (
        '<tr><th>'+escapeHtml(label)+'</th><td>'+escapeHtml(value)+'</td></tr>'
      )).join('') + '</tbody></table>';
      setExpenseDetailMode('view');
    }
    function renderExpenseDetailEdit(index){
      const expense = (data.expenses || [])[index];
      if(!expense || !expenseDetailBody) return;
      const payer = expense.payer || expense.person || '';
      const splitPeople = uniqueList([...(fileData.people || []), payer, ...Object.keys(expense.splits || {})]);
      const currentSplits = Object.keys(expense.splits || {}).length ? expense.splits : {[payer]:100};
      expenseDetailBody.innerHTML = `
        <div class="expense-edit-grid">
          <div class="col">
            <label for="detailDate">Date</label>
            <input id="detailDate" type="date" value="${escapeHtml(expense.date || dateValueFromMonth(monthEl.value))}">
          </div>
          <div class="col">
            <label for="detailPayer">Payer</label>
            <select id="detailPayer"></select>
          </div>
          <div class="col">
            <label for="detailCategory">Category</label>
            <select id="detailCategory"></select>
          </div>
          <div class="col">
            <label for="detailAmount">Amount</label>
            <input id="detailAmount" type="number" inputmode="decimal" step="0.01" min="0" value="${escapeHtml(+expense.amount || 0)}">
          </div>
          <div class="col expense-edit-wide">
            <label for="detailNote">Note</label>
            <textarea id="detailNote">${escapeHtml(expense.note || '')}</textarea>
          </div>
          <div class="col expense-edit-wide">
            <label>Split</label>
            <div class="expense-edit-splits">
              ${splitPeople.map(person => `
                <label class="expense-edit-split">
                  <span>${escapeHtml(person)}</span>
                  <input class="detailSplitInput" type="number" inputmode="numeric" min="0" max="100" step="5" data-person="${escapeHtml(person)}" value="${escapeHtml(currentSplits[person] || 0)}">
                  <span>%</span>
                </label>
              `).join('')}
            </div>
            <div id="detailSplitStatus" class="note"></div>
          </div>
        </div>
      `;
      const payerSelect = document.getElementById('detailPayer');
      const categorySelect = document.getElementById('detailCategory');
      fillSelect(payerSelect, uniqueList([...(fileData.people || []), payer]));
      fillSelect(categorySelect, uniqueList([...(fileData.categories || []), expense.category || '']));
      payerSelect.value = payer;
      categorySelect.value = expense.category || '';
      const updateSplitStatus = () => {
        const total = Array.from(document.querySelectorAll('.detailSplitInput'))
          .reduce((sum, input) => sum + (parseInt(input.value) || 0), 0);
        const status = document.getElementById('detailSplitStatus');
        if(status) status.textContent = `Split total: ${total}%${total === 100 ? '' : ' - must be 100%'}`;
      };
      document.querySelectorAll('.detailSplitInput').forEach(input => input.addEventListener('input', updateSplitStatus));
      updateSplitStatus();
      setExpenseDetailMode('edit');
    }
    function closeExpenseDetail(){
      if(expenseDetailModal) expenseDetailModal.style.display = 'none';
      if(expenseDetailModal) delete expenseDetailModal.dataset.index;
      setExpenseDetailMode('view');
    }
    function editExpenseFromDetail(){
      const index = expenseDetailIndex();
      if(index >= 0) renderExpenseDetailEdit(index);
    }
    function cancelExpenseEdit(){
      const index = expenseDetailIndex();
      if(index >= 0) renderExpenseDetailView(index);
    }
    async function saveExpenseDetailEdits(){
      const index = expenseDetailIndex();
      const expense = data.expenses?.[index];
      if(!expense) return;
      const date = document.getElementById('detailDate')?.value || dateValueFromMonth(monthEl.value);
      const payer = document.getElementById('detailPayer')?.value || '';
      const category = document.getElementById('detailCategory')?.value || '';
      const note = document.getElementById('detailNote')?.value.trim() || '';
      const amount = parseFloat(document.getElementById('detailAmount')?.value || '');
      if(!payer){ alert('Please choose a payer.'); return; }
      if(!category){ alert('Please choose a category.'); return; }
      if(!Number.isFinite(amount) || amount <= 0){ alert('Please enter a valid amount.'); return; }
      const splits = {};
      let splitTotal = 0;
      document.querySelectorAll('.detailSplitInput').forEach(input => {
        const value = parseInt(input.value) || 0;
        if(value > 0) splits[input.dataset.person] = value;
        splitTotal += value;
      });
      if(splitTotal !== 100){
        alert('The split must add up to 100%.');
        return;
      }
      data.expenses[index] = { ...expense, date, payer, category, note, amount: round2(amount), splits };
      delete data.expenses[index].person;
      delete data.expenses[index].dateEstimated;
      await save();
      renderAll();
      renderExpenseDetailView(index);
    }
    async function deleteExpenseFromDetail(){
      const index = expenseDetailIndex();
      if(!Number.isInteger(index) || !data.expenses?.[index]) return;
      if(!confirm('Delete this expense?')) return;
      data.expenses.splice(index, 1);
      await save();
      closeExpenseDetail();
      renderAll();
    }
    function renderHome(){
      if(homeMonthLabel) homeMonthLabel.textContent = monthLabel(monthEl.value);
      const expenses = data?.expenses || [];
      const total = expenses.reduce((sum, expense) => round2(sum + (+expense.amount || 0)), 0);
      if(homeMonthlyTotal) homeMonthlyTotal.textContent = euro(round2(total));
      if(homeEntryCount) homeEntryCount.textContent = String(expenses.length);
    }
    function renderAll(){
      renderPeopleCats(); renderExpenses(); renderTotals(); renderSettlements(); renderHome(); renderLastEdited(); renderAverages();
    }

    function sectionsForView(view){
      if(!currentUser) return ['auth'];
      const map = {
        home: ['home', 'month'],
        add: ['month', 'add-expenses'],
        review: ['month', 'all-expenses'],
        settle: ['month', 'totals-people', 'settlements'],
        analysis: ['month', 'averages'],
        account: ['auth'],
        settings: ['people', 'categories']
      };
      return map[view] || map.home;
    }

    function setView(view){
      activeView = currentUser ? (view || 'home') : 'account';
      if(!currentUser) closeMobileMenu();
      const visible = new Set(sectionsForView(activeView));
      document.querySelectorAll('[data-section]').forEach(section => {
        section.classList.toggle('view-hidden', !visible.has(section.dataset.section));
      });
      if(appNav){
        appNav.style.display = '';
        appNav.classList.toggle('signed-in', !!currentUser);
        appNav.classList.toggle('signed-out', !currentUser);
        appNav.querySelectorAll('[data-target-view]').forEach(button => {
          button.classList.toggle('active', button.dataset.targetView === activeView);
        });
      }
      if(menuToggle){
        menuToggle.hidden = false;
      }
      renderHome();
    }

    function refreshAppView(){
      setView(currentUser ? activeView || 'home' : 'account');
    }
    function closeMobileMenu(){
      if(appNav) appNav.classList.remove('menu-open');
      if(menuToggle) menuToggle.setAttribute('aria-expanded', 'false');
    }
    function toggleMobileMenu(){
      if(!appNav || !menuToggle) return;
      const isOpen = appNav.classList.toggle('menu-open');
      menuToggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    }

    function makeExportSnapshot(){
      const months = {};
      for(const month of listSavedMonths()){
        const key = normalizeMonthKey(month);
        if(/^\d{4}-\d{2}$/.test(key)){
          months[key] = normalizeMonthData(getMonthData(key));
        }
      }
      return {
        exportedAt: new Date().toISOString(),
        mode: currentUser ? 'signed-in' : 'local',
        user: currentUser ? { id: currentUser.id, email: currentUser.email } : null,
        people: fileData.people || [],
        categories: fileData.categories || [],
        months
      };
    }

    function downloadText(filename, text, type = 'application/json'){
      const blob = new Blob([text], { type });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      a.click();
      URL.revokeObjectURL(url);
    }

    function exportAllData(){
      const snapshot = makeExportSnapshot();
      const stamp = new Date().toISOString().slice(0, 10);
      downloadText(`family-accounts-export-${stamp}.json`, JSON.stringify(snapshot, null, 2));
    }

    function renderBackupRows(backups){
      if(!backupList) return;
      backupList.innerHTML = '';
      if(!backups.length){
        backupList.innerHTML = '<div class="muted">No backups found yet.</div>';
        return;
      }
      const table = document.createElement('table');
      table.innerHTML = '<thead><tr><th>When</th><th>Source</th><th>Operation</th><th>Old</th><th>New</th><th></th></tr></thead><tbody></tbody>';
      const tbody = table.querySelector('tbody');
      backups.forEach(backup => {
        const oldPayload = backup.old_data?.data || backup.old_data || null;
        const newPayload = backup.new_data?.data || backup.new_data || null;
        const oldCount = backup.source_table === 'family_account_months' ? expenseCount(normalizeMonthData(oldPayload)) : '';
        const newCount = backup.source_table === 'family_account_months' ? expenseCount(normalizeMonthData(newPayload)) : '';
        const tr = document.createElement('tr');
        [formatUK(backup.created_at), backup.source_month || backup.source_table, backup.operation, oldCount, newCount].forEach(value => {
          const td = document.createElement('td');
          td.textContent = String(value ?? '');
          tr.appendChild(td);
        });
        const actionCell = document.createElement('td');
        if(backup.source_table === 'family_account_months' && backup.source_month){
          const restoreBtn = document.createElement('button');
          restoreBtn.type = 'button';
          restoreBtn.className = 'btn ghost compact';
          restoreBtn.textContent = 'Restore';
          restoreBtn.addEventListener('click', async () => {
            try{
              const restored = await restoreMonthFromBackup(backup);
              if(restored){
                backupStatus.textContent = `Restored ${backup.source_month}.`;
                await loadAndRenderBackups();
              }
            }catch(err){
              backupStatus.textContent = err.message || String(err);
            }
          });
          actionCell.appendChild(restoreBtn);
        }
        tr.appendChild(actionCell);
        tbody.appendChild(tr);
      });
      backupList.appendChild(table);
    }

    async function loadAndRenderBackups(){
      if(!backupStatus || !backupList) return;
      backupStatus.textContent = 'Loading backups...';
      try{
        const backups = await fetchRecentBackups(40);
        renderBackupRows(backups);
        backupStatus.textContent = `Loaded ${backups.length} recent backups.`;
      }catch(err){
        backupStatus.textContent = err.message || String(err);
      }
    }

    function fillSelect(selectEl, items){
      const currentValue = selectEl.value; // Preserve current selection
      selectEl.innerHTML = '';
      const opt = document.createElement('option'); opt.value=''; opt.textContent='— select —'; selectEl.appendChild(opt);
      for(const it of (items||[])){ const o = document.createElement('option'); o.value=it; o.textContent=it; selectEl.appendChild(o); }
      if(currentValue) selectEl.value = currentValue; // Restore selection if it still exists
    }

    function getAvailableYears(months){
      const years = new Set();
      for(const m of months){
        const key = normalizeMonthKey(m);
        if(key && key.includes('-')){
          years.add(key.split('-')[0]);
        }
      }
      return Array.from(years).sort((a,b)=> Number(b) - Number(a));
    }

    // Analysis across months (from local or file)
    function listSavedMonths(){ return currentUser ? listSavedMonthsRemote() : listSavedMonthsLocal(); }
    function getMonthData(m){
      if(currentUser){
        const mapKey = findMonthKeyInMap(m);
        return fileData.months[mapKey] || defaultData(null);
      }
      const raw = localStorage.getItem(keyFor(m)); if(!raw) return defaultData(null);
      try{ return JSON.parse(raw); }catch{return defaultData(null);} }
    function monthLabel(monthKey){
      const key = normalizeMonthKey(monthKey);
      if(!/^\d{4}-\d{2}$/.test(key)) return key;
      const [year, month] = key.split('-');
      return new Date(Number(year), Number(month)-1, 1).toLocaleDateString('en-GB', { month:'short', year:'numeric' });
    }
    function summarizeExpenses(expenses){
      const categoryTotals = {};
      const payerTotals = {};
      let total = 0;
      for(const e of (expenses || [])){
        const amount = +e.amount || 0;
        total = round2(total + amount);
        const category = e.category || 'Uncategorised';
        const payer = e.payer || e.person || 'Unknown';
        categoryTotals[category] = round2((categoryTotals[category] || 0) + amount);
        payerTotals[payer] = round2((payerTotals[payer] || 0) + amount);
      }
      const count = (expenses || []).length;
      return { total: round2(total), count, average: count ? round2(total / count) : 0, categoryTotals, payerTotals };
    }
    function getAnalysisMonthRows(){
      return listSavedMonths()
        .map(m => normalizeMonthKey(m))
        .filter(m => /^\d{4}-\d{2}$/.test(m))
        .filter((m, idx, arr) => arr.indexOf(m) === idx)
        .sort()
        .map(m => ({ month:m, data:getMonthData(m) }))
        .filter(row => row.data && (row.data.expenses || []).length);
    }

    async function migrateMissingExpenseDates(){
      const months = listSavedMonths().map(normalizeMonthKey).filter(m => /^\d{4}-\d{2}$/.test(m));
      let changedExpenseCount = 0;
      const changedMonths = [];
      for(const month of months){
        const monthData = normalizeMonthData(getMonthData(month));
        let changed = false;
        for(const expense of monthData.expenses || []){
          if(!expense.date){
            expense.date = dateValueFromMonth(month);
            expense.dateEstimated = true;
            changed = true;
            changedExpenseCount += 1;
          }
        }
        if(changed){
          monthData.lastEdited = new Date().toISOString();
          changedMonths.push({ month, data: monthData });
        }
      }
      if(!changedMonths.length){
        alert('No undated expenses found.');
        return;
      }
      if(!confirm(`Add estimated dates to ${changedExpenseCount} old expenses across ${changedMonths.length} months?`)){
        return;
      }
      if(currentUser){
        for(const row of changedMonths){
          fileData.months[row.month] = row.data;
          localSaveCurrent(row.month, row.data);
        }
        await saveRemote();
        for(const row of changedMonths){
          await saveRemoteMonth(row.month, row.data, { allowEmptyOverwrite:true });
        }
        await loadFromFileOrLocal(monthEl.value);
      } else {
        changedMonths.forEach(row => localSaveCurrent(row.month, row.data));
        await loadFromFileOrLocal(monthEl.value);
      }
      renderAll();
      alert(`Estimated dates added to ${changedExpenseCount} expenses.`);
    }
    function sumMaps(maps){
      const total = {};
      for(const map of maps){
        for(const [key, value] of Object.entries(map || {})){
          total[key] = round2((total[key] || 0) + (+value || 0));
        }
      }
      return total;
    }
    function tableFromMap(mapObj, leftHeader, rightHeader){
      const keys = Object.keys(mapObj).sort((a,b)=> mapObj[b]-mapObj[a]); if(!keys.length) return '<div class="muted">No data yet.</div>';
      let html = '<table><thead><tr><th>'+leftHeader+'</th><th>'+rightHeader+'</th></tr></thead><tbody>';
      for(const k of keys){ html += '<tr><td>'+k+'</td><td>'+euro(mapObj[k])+'</td></tr>'; }
      html += '</tbody></table>'; return html;
    }
    function tableFromRows(rows, headers){
      if(!rows.length) return '<div class="muted">No data yet.</div>';
      let html = '<table><thead><tr>'+headers.map(h=>'<th>'+h+'</th>').join('')+'</tr></thead><tbody>';
      for(const row of rows){
        html += '<tr>'+row.map(cell=>'<td>'+cell+'</td>').join('')+'</tr>';
      }
      html += '</tbody></table>';
      return html;
    }
    function renderMetricCards(metrics){
      analysisSummary.innerHTML = metrics.map(metric => (
        '<div class="analysis-metric"><div class="muted">'+metric.label+'</div><div class="big">'+metric.value+'</div></div>'
      )).join('');
    }
    function setCompareOptions(options){
      const controls = document.querySelectorAll('.analysis-compare-control');
      const isCompare = analysisView && (analysisView.value === 'compare-months' || analysisView.value === 'compare-years');
      controls.forEach(control => { control.style.display = isCompare ? 'flex' : 'none'; });
      if(!isCompare) return;
      const previousA = analysisCompareA.value;
      const previousB = analysisCompareB.value;
      analysisCompareA.innerHTML = '';
      analysisCompareB.innerHTML = '';
      for(const option of options){
        const a = document.createElement('option');
        a.value = option.value;
        a.textContent = option.label;
        analysisCompareA.appendChild(a);
        const b = document.createElement('option');
        b.value = option.value;
        b.textContent = option.label;
        analysisCompareB.appendChild(b);
      }
      if(options.some(o => o.value === previousA)) analysisCompareA.value = previousA;
      if(options.some(o => o.value === previousB)) analysisCompareB.value = previousB;
      if(!analysisCompareA.value && options[0]) analysisCompareA.value = options[0].value;
      if(!analysisCompareB.value && options[1]) analysisCompareB.value = options[1].value;
      if(analysisCompareA.value === analysisCompareB.value && options.length > 1){
        analysisCompareB.value = options.find(o => o.value !== analysisCompareA.value)?.value || analysisCompareB.value;
      }
    }
    function renderAverages(){
      if(!analysisView || !analysisSummary || !analysisTable) return;
      const monthRows = getAnalysisMonthRows();
      const years = getAvailableYears(monthRows.map(row => row.month));
      const view = analysisView.value || 'month';
      const allSummary = summarizeExpenses(monthRows.flatMap(row => row.data.expenses || []));
      const selectedMonth = normalizeMonthKey(monthEl.value);
      const currentMonth = monthRows.find(row => row.month === selectedMonth) || { month:selectedMonth, data:data || defaultData() };
      const currentSummary = summarizeExpenses(currentMonth.data.expenses || []);

      if(view !== 'compare-months' && view !== 'compare-years') setCompareOptions([]);

      if(view === 'month'){
        renderMetricCards([
          { label: monthLabel(selectedMonth), value: euro(currentSummary.total) },
          { label: 'Entries', value: String(currentSummary.count) },
          { label: 'Average entry', value: euro(currentSummary.average) }
        ]);
        analysisTable.innerHTML = tableFromMap(currentSummary.categoryTotals, 'Category', 'Total');
        return;
      }

      if(view === 'months'){
        const rows = monthRows.slice().reverse().map(row => {
          const summary = summarizeExpenses(row.data.expenses || []);
          return [monthLabel(row.month), euro(summary.total), String(summary.count), euro(summary.average)];
        });
        renderMetricCards([
          { label: 'All months total', value: euro(allSummary.total) },
          { label: 'Months with data', value: String(monthRows.length) },
          { label: 'Average per month', value: euro(monthRows.length ? round2(allSummary.total / monthRows.length) : 0) }
        ]);
        analysisTable.innerHTML = tableFromRows(rows, ['Month', 'Total', 'Entries', 'Avg entry']);
        return;
      }

      const yearGroups = {};
      for(const row of monthRows){
        const year = row.month.split('-')[0];
        if(!yearGroups[year]) yearGroups[year] = [];
        yearGroups[year].push(row);
      }

      if(view === 'years'){
        const rows = Object.keys(yearGroups).sort((a,b)=>Number(b)-Number(a)).map(year => {
          const expenses = yearGroups[year].flatMap(row => row.data.expenses || []);
          const summary = summarizeExpenses(expenses);
          return [year, euro(summary.total), String(yearGroups[year].length), euro(yearGroups[year].length ? round2(summary.total / yearGroups[year].length) : 0)];
        });
        renderMetricCards([
          { label: 'All years total', value: euro(allSummary.total) },
          { label: 'Years with data', value: String(Object.keys(yearGroups).length) },
          { label: 'Average per month', value: euro(monthRows.length ? round2(allSummary.total / monthRows.length) : 0) }
        ]);
        analysisTable.innerHTML = tableFromRows(rows, ['Year', 'Total', 'Months', 'Avg / month']);
        return;
      }

      if(view === 'ytd'){
        const currentYear = String(new Date().getFullYear());
        const year = years.includes(currentYear) ? currentYear : (years[0] || currentYear);
        const cutoffMonth = year === currentYear ? thisMonthValue() : `${year}-12`;
        const ytdRows = monthRows.filter(row => row.month.startsWith(year + '-') && row.month <= cutoffMonth);
        const summary = summarizeExpenses(ytdRows.flatMap(row => row.data.expenses || []));
        renderMetricCards([
          { label: year + ' total', value: euro(summary.total) },
          { label: 'Months included', value: String(ytdRows.length) },
          { label: 'Average per month', value: euro(ytdRows.length ? round2(summary.total / ytdRows.length) : 0) }
        ]);
        analysisTable.innerHTML = tableFromMap(summary.categoryTotals, 'Category', 'YTD total');
        return;
      }

      if(view === 'averages'){
        const categoryAverages = {};
        const categoryTotals = sumMaps(monthRows.map(row => summarizeExpenses(row.data.expenses || []).categoryTotals));
        for(const [category, total] of Object.entries(categoryTotals)){
          categoryAverages[category] = monthRows.length ? round2(total / monthRows.length) : 0;
        }
        renderMetricCards([
          { label: 'Average per month', value: euro(monthRows.length ? round2(allSummary.total / monthRows.length) : 0) },
          { label: 'Average entry', value: euro(allSummary.average) },
          { label: 'Entries analysed', value: String(allSummary.count) }
        ]);
        analysisTable.innerHTML = tableFromMap(categoryAverages, 'Category', 'Avg / month');
        return;
      }

      if(view === 'compare-months'){
        const options = monthRows.slice().reverse().map(row => ({ value:row.month, label:monthLabel(row.month) }));
        setCompareOptions(options);
        const a = monthRows.find(row => row.month === analysisCompareA.value);
        const b = monthRows.find(row => row.month === analysisCompareB.value);
        renderComparison(a?.data?.expenses || [], b?.data?.expenses || [], monthLabel(analysisCompareA.value), monthLabel(analysisCompareB.value));
        return;
      }

      if(view === 'compare-years'){
        const options = Object.keys(yearGroups).sort((a,b)=>Number(b)-Number(a)).map(year => ({ value:year, label:year }));
        setCompareOptions(options);
        const aExpenses = (yearGroups[analysisCompareA.value] || []).flatMap(row => row.data.expenses || []);
        const bExpenses = (yearGroups[analysisCompareB.value] || []).flatMap(row => row.data.expenses || []);
        renderComparison(aExpenses, bExpenses, analysisCompareA.value, analysisCompareB.value);
      }
    }
    function renderComparison(aExpenses, bExpenses, aLabel, bLabel){
      const a = summarizeExpenses(aExpenses);
      const b = summarizeExpenses(bExpenses);
      renderMetricCards([
        { label: aLabel || 'First', value: euro(a.total) },
        { label: bLabel || 'Second', value: euro(b.total) },
        { label: 'Difference', value: euro(round2(a.total - b.total)) }
      ]);
      const categories = Array.from(new Set([...Object.keys(a.categoryTotals), ...Object.keys(b.categoryTotals)])).sort();
      const rows = categories.map(category => {
        const av = a.categoryTotals[category] || 0;
        const bv = b.categoryTotals[category] || 0;
        return [category, euro(av), euro(bv), euro(round2(av - bv))];
      });
      analysisTable.innerHTML = tableFromRows(rows, ['Category', aLabel || 'First', bLabel || 'Second', 'Difference']);
    }

    // ---------- Supabase ----------
    function initSupabase(){
      if(!window.supabase || !SUPABASE_URL || !SUPABASE_ANON_KEY){
        console.warn('Supabase client not configured.');
        return;
      }
      supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY);
    }

    async function loadRemote(){
      if(!supabaseClient || !currentUser) return false;
      remoteLoadState = 'loading';
      updateCloudStatus();
      let needsSave = false;
      try{
        const { data: rows, error } = await supabaseClient
          .from('family_account_data')
          .select('data')
          .eq('user_id', currentUser.id)
          .limit(1);
        if(error) throw error;
        if(rows && rows.length){
          fileData = rows[0].data || { people: [], categories: [], months: {} };
        } else {
          fileData = { people: [], categories: [], months: {} };
          needsSave = true;
        }

        if(!fileData.people) fileData.people = [];
        if(!fileData.categories) fileData.categories = [];
        if(!fileData.months || typeof fileData.months !== 'object') fileData.months = {};

        // Load legacy per-month rows if they exist and merge into months map.
        const monthsRows = await fetchMonthsRows();
        const monthsRowMap = {};
        if(Array.isArray(monthsRows) && monthsRows.length){
          for(const row of monthsRows){
            const key = normalizeMonthKey(row?.month);
            const incoming = row?.data ? normalizeMonthData(row.data) : null;
            if(key && incoming){
              monthsRowMap[key] = incoming;
            }
          }
        }
        for(const [key, incoming] of Object.entries(monthsRowMap)){
          const existing = key ? fileData.months[key] : null;
          if(key && incoming && shouldReplaceMonth(existing, incoming)){
            fileData.months[key] = incoming;
            needsSave = true;
          }
        }

        // Backfill per-month rows from the global months map if they are missing/empty.
        const backfillMonths = [];
        for(const [rawKey, monthData] of Object.entries(fileData.months || {})){
          const key = normalizeMonthKey(rawKey);
          const normalized = normalizeMonthData(monthData);
          const existingRow = monthsRowMap[key];
          if(shouldReplaceMonth(existingRow, normalized)){
            monthsRowMap[key] = normalized;
            backfillMonths.push(key);
          }
          if(rawKey !== key){
            fileData.months[key] = normalized;
            needsSave = true;
          }
        }
        for(const key of backfillMonths){
          await saveRemoteMonth(key, monthsRowMap[key]);
        }

        // Local storage fallback: prefer newer local data if remote is stale.
        const localMonths = listSavedMonthsLocal();
        const localUpdates = [];
        for(const localKey of localMonths){
          const localData = getLocalMonthData(localKey);
          if(!localData) continue;
          const key = normalizeMonthKey(localKey);
          const existing = fileData.months[key];
          if(isNewerMonthData(localData, existing)){
            fileData.months[key] = localData;
            localUpdates.push(key);
            needsSave = true;
          }
        }
        for(const key of localUpdates){
          await saveRemoteMonth(key, fileData.months[key]);
        }

        // Migrate people/categories from month records if missing.
        if(fileData.people.length === 0 && fileData.categories.length === 0){
          const allPeople = new Set();
          const allCategories = new Set();
          for(const monthData of Object.values(fileData.months)){
            if(monthData && typeof monthData === 'object'){
              if(Array.isArray(monthData.people)){
                monthData.people.forEach(p => allPeople.add(String(p).trim()));
              }
              if(Array.isArray(monthData.categories)){
                monthData.categories.forEach(c => allCategories.add(String(c).trim()));
              }
            }
          }
          if(allPeople.size > 0 || allCategories.size > 0){
            fileData.people = Array.from(allPeople).filter(Boolean);
            fileData.categories = Array.from(allCategories).filter(Boolean);
            needsSave = true;
          }
        }

        remoteLoadState = 'ready';
        updateCloudStatus();
        if(needsSave){
          await saveRemote();
        }
        return true;
      }catch(err){
        remoteLoadState = 'failed';
        updateCloudStatus();
        console.error('Failed to load remote data', err);
        return false;
      }
    }

    async function loadRemoteMonth(month){
      return await fetchMonthRow(month);
    }

    async function loadRemoteMonthStrict(month){
      return await fetchMonthRow(month, true);
    }

    async function fetchRemoteFileData(){
      const { data: remoteRows, error: remoteError } = await withTimeout(
        supabaseClient
          .from('family_account_data')
          .select('data')
          .eq('user_id', currentUser.id)
          .limit(1),
        8000,
        'Cloud safety check'
      );
      if(remoteError) throw remoteError;
      return remoteRows?.[0]?.data || {};
    }

    async function fetchRecentBackups(limit = 30){
      if(!supabaseClient || !currentUser) throw new Error('Sign in before loading backups.');
      const { data: rows, error } = await withTimeout(
        supabaseClient
          .from('family_account_backups')
          .select('id,source_table,source_month,operation,old_data,new_data,created_at')
          .order('created_at', { ascending:false })
          .limit(limit),
        8000,
        'Backup load'
      );
      if(error) throw error;
      return rows || [];
    }

    function backupPayloadForRestore(backup){
      if(!backup) return null;
      const row = backup.old_data || backup.new_data || null;
      if(!row) return null;
      if(backup.source_table === 'family_account_months'){
        return row.data ? normalizeMonthData(row.data) : null;
      }
      return row.data || null;
    }

    async function restoreMonthFromBackup(backup){
      if(!backup || backup.source_table !== 'family_account_months' || !backup.source_month){
        throw new Error('Choose a month backup to restore.');
      }
      const restored = backupPayloadForRestore(backup);
      if(!restored || !Array.isArray(restored.expenses)){
        throw new Error('This backup does not contain a restorable month.');
      }
      const month = normalizeMonthKey(backup.source_month);
      const count = expenseCount(restored);
      const message = `Restore ${month} from backup ${backup.id} with ${count} expenses? This will replace the current cloud month.`;
      if(!confirm(message)) return false;
      fileData.months[month] = restored;
      if(month === normalizeMonthKey(activeMonth || monthEl.value)){
        data = JSON.parse(JSON.stringify(restored));
        localSaveCurrent(month, data);
      }
      await saveRemote({ allowEmptyOverwrite:true });
      await saveRemoteMonth(month, restored, { allowEmptyOverwrite:true });
      await loadFromFileOrLocal(monthEl.value);
      renderAll();
      return true;
    }

    async function saveRemoteMonth(month, monthData, options = {}){
      const key = normalizeMonthKey(month);
      const nextData = normalizeMonthData(monthData);
      if(!options.allowEmptyOverwrite && !hasExpenses(nextData)){
        const existing = await loadRemoteMonthStrict(key);
        if(hasExpenses(existing)){
          throw new Error(`Refused to overwrite ${key} in the cloud because the app has an empty month locally but Supabase has ${expenseCount(existing)} expenses.`);
        }
        const remoteFileData = await fetchRemoteFileData();
        const globalExisting = normalizeMonthData(remoteFileData?.months?.[key]);
        if(hasExpenses(globalExisting)){
          throw new Error(`Refused to overwrite ${key} in the cloud because the app has an empty month locally but the main Supabase file has ${expenseCount(globalExisting)} expenses.`);
        }
      }
      await upsertMonthRow(key, nextData);
    }

    async function saveRemote(options = {}){
      if(!supabaseClient || !currentUser) return;
      if(remoteLoadState !== 'ready'){
        throw new Error('Cloud data has not loaded, so upload was blocked.');
      }
      // Capture fileData snapshot immediately to prevent corruption if fileData is reset during save
      const dataSnapshot = JSON.parse(JSON.stringify(fileData));
      const remoteData = await fetchRemoteFileData();
      storeCloudSafetyBackup(remoteData);
      dataSnapshot.months = mergeProtectedMonths(remoteData.months || {}, dataSnapshot.months || {}, options);
      const payload = {
        user_id: currentUser.id,
        data: dataSnapshot,
        updated_at: new Date().toISOString()
      };
      const { error } = await withTimeout(
        supabaseClient
          .from('family_account_data')
          .upsert(payload, { onConflict: 'user_id' }),
        8000,
        'Cloud save'
      );
      if(error) throw error;
    }

    async function refreshSession(){
      if(!supabaseClient) return;
      const { data: sessionData } = await supabaseClient.auth.getSession();
      if(sessionData?.session){
        currentSession = sessionData.session;
        currentUser = sessionData.session.user;
      }
      setModeBadge();
    }

    function updateAuthStatus(){
      if(currentUser){
        authStatus.textContent = `Signed in as ${currentUser.email}`;
        if(accountCardTitle) accountCardTitle.textContent = 'Account management';
        if(accountSectionTitle) accountSectionTitle.textContent = 'Account';
      } else {
        authStatus.textContent = 'Not signed in.';
        if(accountCardTitle) accountCardTitle.textContent = 'Sign in';
        if(accountSectionTitle) accountSectionTitle.textContent = 'Sign in';
      }
      const authFieldsRow = authEmail?.closest('.row');
      if(authFieldsRow) authFieldsRow.style.display = currentUser ? 'none' : 'flex';
      [signInBtn, signUpBtn, forgotPasswordBtn].forEach(button => {
        if(button) button.style.display = currentUser ? 'none' : '';
      });
      [signOutBtn, saveCloudBtn, retryCloudBtn, exportDataBtn].forEach(button => {
        if(button) button.style.display = currentUser ? '' : 'none';
      });
      document.querySelectorAll('.signed-in-only').forEach(section => {
        section.style.display = currentUser ? '' : 'none';
      });
      if(mobileMenuSignInBtn) mobileMenuSignInBtn.style.display = currentUser ? 'none' : '';
      if(mobileMenuSignOutBtn) mobileMenuSignOutBtn.style.display = currentUser ? '' : 'none';
      const advancedTools = document.querySelector('.advanced-tools');
      if(advancedTools) advancedTools.style.display = currentUser ? 'block' : 'none';
      setModeBadge();
      updateCloudStatus();
    }

    // ---------- Month switching ----------
    async function switchMonthSafely(targetMonth){
      const normalizedMonth = normalizeMonthKey(targetMonth);
      if(normalizedMonth === lastLoadedMonth) return;
      await loadFromFileOrLocal(normalizedMonth);
      if(!data){
        data = defaultData();
      }
      lastLoadedMonth = normalizedMonth;
      monthEl.value = normalizedMonth;
      renderLastEdited();
      renderAll();
      renderAverages();
    }

    // ---------- Public actions ----------
    function setMonthToNow(){ monthEl.value = thisMonthValue(); }
    function loadMonth(month){ loadFromFileOrLocal(month).then(()=>{ renderAll(); renderAverages(); }); }
    function syncEntryDatesToMonth(){
      const fallbackDate = dateValueFromMonth(monthEl.value);
      if(quickDate && (!quickDate.value || !quickDate.value.startsWith(normalizeMonthKey(monthEl.value)))) quickDate.value = fallbackDate;
      if(batchDate && (!batchDate.value || !batchDate.value.startsWith(normalizeMonthKey(monthEl.value)))) batchDate.value = fallbackDate;
    }

    async function save(options = {}){ setLastEditedNow(); await persist(options); }
    async function addPerson(name){
      name = (name||'').trim(); if(!name) return; if((fileData.people||[]).includes(name)) return;
      (fileData.people = fileData.people||[]).push(name);
      saveGlobalToLocal();
      if(currentUser){
        try{
          await saveRemote();
        }catch(err){
          console.error('Failed to sync person to cloud:', err);
          alert('Warning: Failed to sync to cloud. Please check your connection.');
        }
      }
      renderPeopleCats();
    }
    async function removePerson(name){
      // Remove from global list
      fileData.people = (fileData.people||[]).filter(p=>p!==name);
      saveGlobalToLocal();

      // Remove from current month expenses
      data.expenses = (data.expenses||[]).filter(e=>e.payer!==name && e.person!==name);

      // If using Supabase, remove from all months
      if(currentUser && fileData.months){
        for(const month in fileData.months){
          if(fileData.months[month] && fileData.months[month].expenses){
            fileData.months[month].expenses = fileData.months[month].expenses.filter(e=>e.payer!==name && e.person!==name);
          }
        }
        try{
          await saveRemote();
        }catch(err){
          console.error('Failed to sync person removal to cloud:', err);
          alert('Warning: Failed to sync to cloud. Please check your connection.');
        }
      } else {
        // If using local storage, clean up all saved months
        for(const month of listSavedMonthsLocal()){
          const raw = localStorage.getItem(keyFor(month));
          if(raw){
            try{
              const monthData = JSON.parse(raw);
              if(monthData.expenses){
                monthData.expenses = monthData.expenses.filter(e=>e.payer!==name && e.person!==name);
                localStorage.setItem(keyFor(month), JSON.stringify(monthData));
              }
            }catch{}
          }
        }
      }

      await save(); renderAll();
    }
    async function addCategory(name){
      name = (name||'').trim(); if(!name) return; if((fileData.categories||[]).includes(name)) return;
      (fileData.categories = fileData.categories||[]).push(name);
      saveGlobalToLocal();
      if(currentUser){
        try{
          await saveRemote();
        }catch(err){
          console.error('Failed to sync category to cloud:', err);
          alert('Warning: Failed to sync to cloud. Please check your connection.');
        }
      }
      renderPeopleCats();
    }
    function renameCategoryInMonth(monthData, oldName, newName){
      if(!monthData || typeof monthData !== 'object') return false;
      let changed = false;
      if(Array.isArray(monthData.categories)){
        monthData.categories = uniqueList(monthData.categories.map(c => c === oldName ? newName : c));
        changed = true;
      }
      if(Array.isArray(monthData.expenses)){
        for(const expense of monthData.expenses){
          if(expense?.category === oldName){
            expense.category = newName;
            changed = true;
          }
        }
      }
      if(changed){
        monthData.lastEdited = new Date().toISOString();
      }
      return changed;
    }
    async function renameCategory(oldName, newName){
      oldName = String(oldName || '').trim();
      newName = String(newName || '').trim();
      if(!oldName || !newName || oldName === newName) return;
      if((fileData.categories || []).includes(newName)){
        alert('That category already exists.');
        return;
      }

      fileData.categories = uniqueList((fileData.categories || []).map(c => c === oldName ? newName : c));
      saveGlobalToLocal();

      renameCategoryInMonth(data, oldName, newName);

      if(currentUser && fileData.months){
        const changedMonths = [];
        for(const month of Object.keys(fileData.months)){
          if(renameCategoryInMonth(fileData.months[month], oldName, newName)){
            changedMonths.push({ month: normalizeMonthKey(month), data: fileData.months[month] });
          }
        }
        if(activeMonth){
          fileData.months[normalizeMonthKey(activeMonth)] = normalizeMonthData(data);
        }
        try{
          await saveRemote();
          for(const changedMonth of changedMonths){
            await saveRemoteMonth(changedMonth.month, changedMonth.data);
          }
        }catch(err){
          console.error('Failed to sync category rename to cloud:', err);
          alert('Warning: Failed to sync to cloud. Please check your connection.');
        }
      } else {
        for(const month of listSavedMonthsLocal()){
          if(!/^\d{4}-\d{1,2}$/.test(month)) continue;
          const raw = localStorage.getItem(keyFor(month));
          if(raw){
            try{
              const monthData = normalizeMonthData(JSON.parse(raw));
              if(renameCategoryInMonth(monthData, oldName, newName)){
                localStorage.setItem(keyFor(month), JSON.stringify(monthData));
              }
            }catch{}
          }
        }
      }

      await save();
      renderAll();
      renderAverages();
    }
    async function renameCategoryPrompt(oldName){
      const newName = prompt('Rename category:', oldName);
      if(newName === null) return;
      await renameCategory(oldName, newName);
    }
    async function removeCategory(name){
      // Remove from global list
      fileData.categories = (fileData.categories||[]).filter(c=>c!==name);
      saveGlobalToLocal();

      // Remove from current month expenses
      data.expenses = (data.expenses||[]).filter(e=>e.category!==name);

      // If using Supabase, remove from all months
      if(currentUser && fileData.months){
        for(const month in fileData.months){
          if(fileData.months[month] && fileData.months[month].expenses){
            fileData.months[month].expenses = fileData.months[month].expenses.filter(e=>e.category!==name);
          }
        }
        try{
          await saveRemote();
        }catch(err){
          console.error('Failed to sync category removal to cloud:', err);
          alert('Warning: Failed to sync to cloud. Please check your connection.');
        }
      } else {
        // If using local storage, clean up all saved months
        for(const month of listSavedMonthsLocal()){
          const raw = localStorage.getItem(keyFor(month));
          if(raw){
            try{
              const monthData = JSON.parse(raw);
              if(monthData.expenses){
                monthData.expenses = monthData.expenses.filter(e=>e.category!==name);
                localStorage.setItem(keyFor(month), JSON.stringify(monthData));
              }
            }catch{}
          }
        }
      }

      await save(); renderAll();
    }
    function parseAmounts(text){
      const tokens = String(text).split(/[\s,;]+/).map(t=>t.trim()).filter(Boolean);
      const nums = [];
      for(const t of tokens){ const cleaned=t.replace(/\u20ac/g,'').replace(/,/g,'.'); const n=parseFloat(cleaned); if(!isNaN(n)) nums.push(n); }
      return nums;
    }
    function duplicateExpenseKey(expense){
      const payer = String(expense?.payer || expense?.person || '').trim().toLowerCase();
      const category = String(expense?.category || '').trim().toLowerCase();
      const note = String(expense?.note || '').trim().toLowerCase();
      const amount = round2(+expense?.amount || 0).toFixed(2);
      const date = String(expense?.date || '').trim();
      const splits = Object.entries(expense?.splits || {})
        .sort(([a], [b]) => a.localeCompare(b))
        .map(([person, pct]) => `${person}:${Number(pct) || 0}`)
        .join('|');
      return [date, payer, category, note, amount, splits].join('::');
    }
    function importRowToExpense(row){
      const selectedMonth = normalizeMonthKey(monthEl.value);
      return {
        date: row.dateValue || dateValueFromMonth(selectedMonth),
        payer: row.payer,
        category: row.category || 'Other',
        amount: round2(row.amount),
        note: row.description,
        splits: splitFromImportChoice(row.splitChoice, row.payer)
      };
    }
    function findImportDuplicate(row){
      const candidateKey = duplicateExpenseKey(importRowToExpense(row));
      return (data.expenses || []).findIndex(expense => duplicateExpenseKey(expense) === candidateKey);
    }
    function getDuplicateGroups(){
      const expenses = data?.expenses || [];
      const groups = new Map();
      expenses.forEach((expense, index) => {
        const key = duplicateExpenseKey(expense);
        if(!groups.has(key)) groups.set(key, []);
        groups.get(key).push(index);
      });
      return Array.from(groups.values()).filter(indexes => indexes.length > 1);
    }
    function closeDuplicateReview(message){
      duplicateReviewState = null;
      if(duplicateReview){
        duplicateReview.style.display = 'none';
        duplicateReview.innerHTML = '';
      }
      if(message) alert(message);
    }
    function renderDuplicateReview(){
      if(!duplicateReview || !duplicateReviewState) return;
      const { groups, groupIndex, selected } = duplicateReviewState;
      if(groupIndex >= groups.length){
        closeDuplicateReview('Duplicate review complete.');
        return;
      }
      const expenses = data?.expenses || [];
      const indexes = groups[groupIndex].filter(index => expenses[index]);
      if(indexes.length < 2){
        duplicateReviewState.groupIndex += 1;
        renderDuplicateReview();
        return;
      }

      duplicateReview.style.display = 'block';
      duplicateReview.innerHTML = '';

      const first = expenses[indexes[0]];
      const title = document.createElement('div');
      title.className = 'duplicate-review-title';
      title.textContent = `Duplicate group ${groupIndex + 1} of ${groups.length}`;
      duplicateReview.appendChild(title);

      const summary = document.createElement('div');
      summary.className = 'muted';
      summary.textContent = `${first.payer || first.person || 'Unknown payer'} / ${first.category || 'No category'} / ${euro(+first.amount || 0)}${first.note ? ` / ${first.note}` : ''}`;
      duplicateReview.appendChild(summary);

      const tableWrap = document.createElement('div');
      tableWrap.style.overflow = 'auto';
      const table = document.createElement('table');
      table.className = 'duplicate-table';
      table.innerHTML = '<thead><tr><th>Delete?</th><th>#</th><th>Date</th><th>Payer</th><th>Category</th><th>Note</th><th>Amount</th><th>Split</th></tr></thead><tbody></tbody>';
      const tbody = table.querySelector('tbody');
      indexes.forEach((expenseIndex, position) => {
        const expense = expenses[expenseIndex];
        if(!selected.has(expenseIndex) && position > 0){
          selected.add(expenseIndex);
        }
        const payer = expense.payer || expense.person || '';
        const row = document.createElement('tr');
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.checked = selected.has(expenseIndex);
        checkbox.addEventListener('change', () => {
          if(checkbox.checked) selected.add(expenseIndex);
          else selected.delete(expenseIndex);
        });
        const deleteCell = document.createElement('td');
        deleteCell.appendChild(checkbox);
        row.appendChild(deleteCell);
        [expenseIndex + 1, formatExpenseDate(expense.date), payer, expense.category || '', expense.note || '', euro(+expense.amount || 0), splitsToLabel(expense.splits || {[payer]:100})].forEach(value => {
          const td = document.createElement('td');
          td.textContent = value;
          row.appendChild(td);
        });
        tbody.appendChild(row);
      });
      tableWrap.appendChild(table);
      duplicateReview.appendChild(tableWrap);

      const actions = document.createElement('div');
      actions.className = 'duplicate-actions';
      const deleteSelectedBtn = document.createElement('button');
      deleteSelectedBtn.className = 'btn warn';
      deleteSelectedBtn.type = 'button';
      deleteSelectedBtn.textContent = 'Delete selected';
      deleteSelectedBtn.addEventListener('click', applyDuplicateDeletion);
      const keepAllBtn = document.createElement('button');
      keepAllBtn.className = 'btn ghost';
      keepAllBtn.type = 'button';
      keepAllBtn.textContent = 'Keep all / next';
      keepAllBtn.addEventListener('click', () => {
        indexes.forEach(index => selected.delete(index));
        duplicateReviewState.groupIndex += 1;
        renderDuplicateReview();
      });
      const cancelBtn = document.createElement('button');
      cancelBtn.className = 'btn ghost';
      cancelBtn.type = 'button';
      cancelBtn.textContent = 'Cancel review';
      cancelBtn.addEventListener('click', () => closeDuplicateReview());
      actions.appendChild(deleteSelectedBtn);
      actions.appendChild(keepAllBtn);
      actions.appendChild(cancelBtn);
      duplicateReview.appendChild(actions);
    }
    async function applyDuplicateDeletion(){
      if(!duplicateReviewState || duplicateDeleteInProgress) return;
      const currentGroup = duplicateReviewState.groups[duplicateReviewState.groupIndex] || [];
      const selectedInGroup = currentGroup.filter(index => duplicateReviewState.selected.has(index));
      if(!selectedInGroup.length){
        duplicateReviewState.groupIndex += 1;
        renderDuplicateReview();
        return;
      }
      duplicateDeleteInProgress = true;
      try{
        data.expenses = (data.expenses || []).filter((_, index) => !selectedInGroup.includes(index));
        await save();
        renderAll();
        const remainingGroups = getDuplicateGroups();
        if(!remainingGroups.length){
          closeDuplicateReview('Selected duplicates deleted. No duplicate groups remain.');
          return;
        }
        duplicateReviewState = { groups:remainingGroups, groupIndex:0, selected:new Set() };
        renderDuplicateReview();
      } finally {
        duplicateDeleteInProgress = false;
      }
    }
    async function checkDuplicateExpenses(){
      const expenses = data?.expenses || [];
      if(expenses.length < 2){
        closeDuplicateReview('There are not enough expenses in this month to check for duplicates.');
        return;
      }
      const groups = getDuplicateGroups();
      if(!groups.length){
        closeDuplicateReview('No duplicate entries found for this month.');
        return;
      }
      duplicateReviewState = { groups, groupIndex:0, selected:new Set() };
      renderDuplicateReview();
    }

    function normalizeText(value){
      return String(value || '')
        .toLowerCase()
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-z0-9]+/g, ' ')
        .trim();
    }
    function parseImportAmount(value){
      if(typeof value === 'number') return round2(Math.abs(value));
      let cleaned = String(value || '')
        .replace(/\u20ac/g, '')
        .replace(/\s/g, '')
        .replace(/[^\d.-]/g, '');
      const raw = String(value || '').replace(/\u20ac/g, '').replace(/\s/g, '');
      const lastComma = raw.lastIndexOf(',');
      const lastDot = raw.lastIndexOf('.');
      if(lastComma >= 0 && lastDot >= 0){
        cleaned = lastComma > lastDot
          ? raw.replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '')
          : raw.replace(/,/g, '').replace(/[^\d.-]/g, '');
      } else if(lastComma >= 0){
        cleaned = raw.replace(',', '.').replace(/[^\d.-]/g, '');
      } else {
        cleaned = raw.replace(/[^\d.-]/g, '');
      }
      const parsed = parseFloat(cleaned);
      return Number.isFinite(parsed) ? round2(Math.abs(parsed)) : 0;
    }
    function parseImportSignedAmount(value){
      if(typeof value === 'number') return round2(value);
      const raw = String(value || '').replace(/\u20ac/g, '').replace(/\s/g, '');
      const lastComma = raw.lastIndexOf(',');
      const lastDot = raw.lastIndexOf('.');
      let cleaned;
      if(lastComma >= 0 && lastDot >= 0){
        cleaned = lastComma > lastDot
          ? raw.replace(/\./g, '').replace(',', '.').replace(/[^\d.-]/g, '')
          : raw.replace(/,/g, '').replace(/[^\d.-]/g, '');
      } else if(lastComma >= 0){
        cleaned = raw.replace(',', '.').replace(/[^\d.-]/g, '');
      } else {
        cleaned = raw.replace(/[^\d.-]/g, '');
      }
      const parsed = parseFloat(cleaned);
      return Number.isFinite(parsed) ? round2(parsed) : 0;
    }
    function parseImportDate(value, preferMonthFirst = false){
      if(value instanceof Date && !isNaN(value)) return value;
      if(typeof value === 'number' && window.XLSX?.SSF){
        const parsed = XLSX.SSF.parse_date_code(value);
        if(parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
      }
      const text = String(value || '').trim();
      if(!text) return null;
      const iso = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/.exec(text);
      if(iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
      const local = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})/.exec(text);
      if(local){
        const year = Number(local[3].length === 2 ? '20' + local[3] : local[3]);
        const first = Number(local[1]);
        const second = Number(local[2]);
        const monthFirst = preferMonthFirst && first <= 12;
        if(monthFirst || second > 12){
          return new Date(year, first - 1, second);
        }
        return new Date(year, second - 1, first);
      }
      const parsed = new Date(text);
      return isNaN(parsed) ? null : parsed;
    }
    function importDateToMonth(value, preferMonthFirst = false){
      const date = parseImportDate(value, preferMonthFirst);
      if(!date) return '';
      return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
    }
    function getCategoryHistory(){
      const history = {};
      const rows = getAnalysisMonthRows();
      for(const row of rows){
        for(const expense of (row.data.expenses || [])){
          const category = expense.category || '';
          if(!category) continue;
          const noteKey = normalizeText(expense.note || expense.description || '');
          if(noteKey) history[noteKey] = category;
        }
      }
      return history;
    }
    function suggestCategory(description){
      const existing = fileData.categories || [];
      const normalizedDescription = normalizeText(description);
      if(!normalizedDescription) return { category:'', isNew:false, reason:'No description' };

      const history = getCategoryHistory();
      for(const [key, category] of Object.entries(history)){
        if(key && (normalizedDescription.includes(key) || key.includes(normalizedDescription))){
          return { category, isNew:false, reason:'Matched previous expense' };
        }
      }

      for(const category of existing){
        const normalizedCategory = normalizeText(category);
        if(normalizedCategory && normalizedDescription.includes(normalizedCategory)){
          return { category, isNew:false, reason:'Matched existing category name' };
        }
      }

      const keywordRules = [
        { keywords:['mercadona','aldi','lidl','supermercado','carrefour','dia','consum'], labels:['supermercado','groceries','food','provisiones'], fallback:'Groceries' },
        { keywords:['farmacia','pharmacy'], labels:['farmacia','health','medical'], fallback:'Health' },
        { keywords:['netflix','spotify','disney','hbo','prime'], labels:['subscription','subscriptions','streaming'], fallback:'Subscriptions' },
        { keywords:['repsol','cepsa','gasolina','fuel','uber','cabify','taxi','metro','bus'], labels:['transport','travel','fuel'], fallback:'Transport' },
        { keywords:['restaurante','restaurant','bar','cafe','cafeteria'], labels:['restaurant','restaurants','eating out','comer fuera'], fallback:'Restaurants' }
      ];
      for(const rule of keywordRules){
        if(rule.keywords.some(keyword => normalizedDescription.includes(keyword))){
          const existingMatch = existing.find(category => {
            const normalizedCategory = normalizeText(category);
            return rule.labels.some(label => normalizedCategory.includes(label));
          });
          return { category: existingMatch || rule.fallback, isNew: !existingMatch, reason: existingMatch ? 'Matched existing category by keyword' : 'Suggested new category by keyword' };
        }
      }

      return { category:'Other', isNew: !existing.includes('Other'), reason:'No close match' };
    }
    function guessImportColumn(headers, names){
      const normalizedNames = names.map(normalizeText);
      return headers.find(header => {
        const normalizedHeader = normalizeText(header);
        return normalizedNames.some(name => normalizedHeader.includes(name));
      }) || '';
    }
    function importColumnLabel(header){
      const labels = {
        A: 'A - Date',
        B: 'B - Bank category',
        C: 'C - Bank subcategory',
        D: 'D - Payee/details',
        E: 'E - Comment',
        F: 'F - Amount',
        G: 'G - Balance'
      };
      return labels[header] || header;
    }
    function fillColumnSelect(selectEl, headers, selected){
      if(!selectEl) return;
      selectEl.innerHTML = '';
      const none = document.createElement('option');
      none.value = '';
      none.textContent = '— none —';
      selectEl.appendChild(none);
      for(const header of headers){
        const option = document.createElement('option');
        option.value = header;
        option.textContent = importColumnLabel(header);
        selectEl.appendChild(option);
      }
      selectEl.value = selected || '';
    }
    function parseCsv(text){
      const rows = [];
      let row = [], cell = '', inQuotes = false;
      for(let i = 0; i < text.length; i++){
        const char = text[i];
        const next = text[i + 1];
        if(char === '"' && inQuotes && next === '"'){ cell += '"'; i++; continue; }
        if(char === '"'){ inQuotes = !inQuotes; continue; }
        if(char === ',' && !inQuotes){ row.push(cell); cell = ''; continue; }
        if((char === '\n' || char === '\r') && !inQuotes){
          if(char === '\r' && next === '\n') i++;
          row.push(cell); cell = '';
          if(row.some(value => String(value).trim())) rows.push(row);
          row = [];
          continue;
        }
        cell += char;
      }
      row.push(cell);
      if(row.some(value => String(value).trim())) rows.push(row);
      const headers = (rows.shift() || []).map(header => String(header || '').trim());
      return rows.map(values => Object.fromEntries(headers.map((header, index) => [header, values[index] || ''])));
    }
    function excelColumnName(index){
      let name = '';
      let n = index + 1;
      while(n > 0){
        const remainder = (n - 1) % 26;
        name = String.fromCharCode(65 + remainder) + name;
        n = Math.floor((n - 1) / 26);
      }
      return name;
    }
    function rowsToExcelColumnObjects(rows){
      return rows.map(row => {
        const obj = {};
        for(let index = 0; index < 8; index++){
          obj[excelColumnName(index)] = row[index] ?? '';
        }
        return obj;
      });
    }
    function detectImportFormat(rows){
      const hasIngHeader = rows.some(row =>
        normalizeText(row.A).includes('f valor') &&
        normalizeText(row.B).includes('categoria') &&
        normalizeText(row.F).includes('importe')
      );
      return hasIngHeader ? 'ing-direct' : 'generic';
    }
    async function readImportFile(file){
      const extension = file.name.split('.').pop().toLowerCase();
      if(extension === 'csv'){
        return parseCsv(await file.text());
      }
      if(!window.XLSX){
        alert('Spreadsheet reader failed to load. Please check your internet connection and try again.');
        return [];
      }
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type:'array', cellDates:true });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', blankrows:false, raw:false });
      return rowsToExcelColumnObjects(rows);
    }
    async function handleImportFileChange(){
      const file = importFile?.files?.[0];
      if(!file) return;
      importRowsRaw = await readImportFile(file);
      importDraftRows = [];
      if(!importRowsRaw.length){
        alert('No rows found in that spreadsheet.');
        return;
      }
      const headers = Object.keys(importRowsRaw[0] || {});
      importDetectedFormat = detectImportFormat(importRowsRaw);
      const isIngDirect = importDetectedFormat === 'ing-direct';
      fillColumnSelect(importDateColumn, headers, isIngDirect ? 'A' : guessImportColumn(headers, ['date','fecha','transaction date']));
      fillColumnSelect(importDescriptionColumn, headers, isIngDirect ? 'D' : guessImportColumn(headers, ['description','concepto','merchant','comercio','note','details']));
      fillColumnSelect(importAmountColumn, headers, isIngDirect ? 'F' : guessImportColumn(headers, ['amount','importe','value','total']));
      fillColumnSelect(importPayerColumn, headers, isIngDirect ? '' : guessImportColumn(headers, ['payer','person','paid by','pagador']));
      fillSelect(importSamePayer, fileData.people || []);
      if(importSamePayer){
        importSamePayer.value = batchPerson.value || (fileData.people || [])[0] || '';
      }
      if(importUseSamePayer){
        importUseSamePayer.checked = true;
      }
      if(importFormatNotice){
        importFormatNotice.textContent = isIngDirect
          ? 'Detected ING Direct movement export. The preview will use A as date, D/E as note, F as amount, and B/C to help suggest categories.'
          : 'Generic spreadsheet mode. Choose the columns that match this bank export before previewing.';
      }
      if(importMapping) importMapping.style.display = 'block';
      if(importPreview){ importPreview.style.display = 'none'; importPreview.innerHTML = ''; }
    }
    function defaultSplitValue(){
      return 'equal';
    }
    function splitFromImportChoice(choice, payer){
      const people = fileData.people || [];
      if(choice === 'payer') return {[payer]:100};
      if(!people.length) return payer ? {[payer]:100} : {};
      const base = Math.floor(100 / people.length);
      let remaining = 100;
      const splits = {};
      people.forEach((person, index) => {
        const value = index === people.length - 1 ? remaining : base;
        splits[person] = value;
        remaining -= value;
      });
      return splits;
    }
    function updateCategorySuggestionsDatalist(){
      if(!categorySuggestions) return;
      categorySuggestions.innerHTML = '';
      for(const category of (fileData.categories || [])){
        const option = document.createElement('option');
        option.value = category;
        categorySuggestions.appendChild(option);
      }
    }
    function buildImportDraftRows(){
      const selectedMonth = normalizeMonthKey(monthEl.value);
      const dateCol = importDateColumn.value;
      const descriptionCol = importDescriptionColumn.value;
      const amountCol = importAmountColumn.value;
      const payerCol = importPayerColumn.value;
      const isIngDirect = importDetectedFormat === 'ing-direct';
      if(!descriptionCol || !amountCol){
        alert('Please choose at least a description column and an amount column.');
        return;
      }
      const useSamePayer = !!importUseSamePayer?.checked;
      const defaultPayer = importSamePayer?.value || batchPerson.value || (fileData.people || [])[0] || '';
      importDraftRows = importRowsRaw.map((row, sourceIndex) => {
        const parsedDateValue = dateCol ? dateInputValueFromImport(row[dateCol], isIngDirect) : '';
        const rowMonth = parsedDateValue ? parsedDateValue.slice(0, 7) : selectedMonth;
        const rawDescription = String(row[descriptionCol] || '').trim();
        const bankCategoryText = isIngDirect ? [row.B, row.C].map(value => String(value || '').trim()).filter(Boolean).join(' / ') : '';
        const commentText = isIngDirect ? String(row.E || '').trim() : '';
        const description = isIngDirect ? [rawDescription, commentText].filter(Boolean).join(' / ') : rawDescription;
        const suggestionText = isIngDirect ? [bankCategoryText, rawDescription, commentText].filter(Boolean).join(' / ') : rawDescription;
        const signedAmount = parseImportSignedAmount(row[amountCol]);
        const amount = Math.abs(signedAmount);
        const payer = useSamePayer ? defaultPayer : (payerCol ? String(row[payerCol] || '').trim() : defaultPayer);
        const suggestion = suggestCategory(suggestionText);
        const draft = {
          sourceIndex,
          selected: !!description && amount > 0 && rowMonth === selectedMonth,
          date: dateCol ? String(row[dateCol] || '') : '',
          dateValue: parsedDateValue || dateValueFromMonth(selectedMonth),
          month: rowMonth || selectedMonth,
          signedAmount,
          payer,
          description,
          amount,
          category: suggestion.category,
          categoryIsNew: suggestion.isNew,
          reason: suggestion.reason,
          splitChoice: defaultSplitValue()
        };
        draft.duplicateIndex = draft.selected ? findImportDuplicate(draft) : -1;
        if(draft.duplicateIndex >= 0) draft.selected = false;
        return draft;
      }).filter(row => {
        if(!row.description || row.amount <= 0) return false;
        if(row.signedAmount >= 0) return false;
        if(dateCol && !row.month) return false;
        if(isIngDirect && normalizeText(row.date).includes('f valor')) return false;
        if(row.month !== selectedMonth) return false;
        return true;
      });
      renderImportPreview();
    }
    function renderImportStatusCell(cell, row){
      if(!cell) return;
      row.duplicateIndex = findImportDuplicate(row);
      if(row.duplicateIndex >= 0){
        cell.textContent = `Possible duplicate of row ${row.duplicateIndex + 1} already in this month.`;
        cell.classList.add('warning');
        return;
      }
      cell.classList.remove('warning');
      cell.textContent = row.categoryIsNew ? `New category needs confirmation: ${row.category || 'blank'}` : (row.reason || '');
    }
    function renderImportPreview(){
      if(!importPreview) return;
      updateCategorySuggestionsDatalist();
      importPreview.style.display = 'block';
      importPreview.innerHTML = '';
      if(!importDraftRows.length){
        importPreview.innerHTML = '<div class="muted">No rows to preview.</div>';
        return;
      }
      const selectedMonth = normalizeMonthKey(monthEl.value);
      const title = document.createElement('div');
      title.className = 'import-preview-title';
      title.textContent = `Preview for ${monthLabel(selectedMonth)}`;
      importPreview.appendChild(title);

      const list = document.createElement('div');
      list.className = 'import-card-list';
      importDraftRows.forEach((row, index) => {
        const card = document.createElement('div');
        card.className = 'import-card';
        if(row.month !== selectedMonth) card.classList.add('import-row-muted');

        const top = document.createElement('div');
        top.className = 'import-card-top';

        const selected = document.createElement('input');
        selected.type = 'checkbox';
        selected.checked = row.selected;
        selected.addEventListener('change', () => { row.selected = selected.checked; });

        const selectLabel = document.createElement('label');
        selectLabel.className = 'import-select-row';
        selectLabel.appendChild(selected);
        const dateText = document.createElement('span');
        dateText.textContent = row.date || '';
        selectLabel.appendChild(dateText);
        top.appendChild(selectLabel);

        const amountInput = document.createElement('input');
        amountInput.type = 'number';
        amountInput.step = '0.01';
        amountInput.value = row.amount;
        amountInput.addEventListener('input', () => {
          row.amount = parseImportAmount(amountInput.value);
          renderImportStatusCell(statusCell, row);
        });
        top.appendChild(amountInput);
        card.appendChild(top);

        const description = document.createElement('div');
        description.className = 'import-card-description';
        description.textContent = row.description;
        card.appendChild(description);

        const fields = document.createElement('div');
        fields.className = 'import-card-fields';

        const payerField = document.createElement('div');
        payerField.className = 'col';
        const payerLabel = document.createElement('label');
        payerLabel.textContent = 'Payer';
        payerField.appendChild(payerLabel);

        const payerSelect = document.createElement('select');
        fillSelect(payerSelect, fileData.people || []);
        payerSelect.value = row.payer;
        payerSelect.disabled = !!importUseSamePayer?.checked;
        payerSelect.addEventListener('change', () => {
          row.payer = payerSelect.value;
          renderImportStatusCell(statusCell, row);
        });
        payerField.appendChild(payerSelect);
        fields.appendChild(payerField);

        const categoryField = document.createElement('div');
        categoryField.className = 'col';
        const categoryLabel = document.createElement('label');
        categoryLabel.textContent = 'Category';
        categoryField.appendChild(categoryLabel);

        const categorySelect = document.createElement('select');
        fillSelect(categorySelect, fileData.categories || []);
        const existingCategories = fileData.categories || [];
        const suggestedIsNew = row.category && !existingCategories.includes(row.category);
        if(suggestedIsNew){
          const suggestedOption = document.createElement('option');
          suggestedOption.value = row.category;
          suggestedOption.textContent = `New: ${row.category}`;
          categorySelect.appendChild(suggestedOption);
        }
        const customOption = document.createElement('option');
        customOption.value = '__new__';
        customOption.textContent = '+ New category...';
        categorySelect.appendChild(customOption);
        categorySelect.value = row.category || '';
        const newCategoryInput = document.createElement('input');
        newCategoryInput.type = 'text';
        newCategoryInput.placeholder = 'New category';
        newCategoryInput.style.display = 'none';
        const syncCategory = () => {
          if(categorySelect.value === '__new__'){
            newCategoryInput.style.display = 'block';
            row.category = newCategoryInput.value.trim();
          } else {
            newCategoryInput.style.display = 'none';
            row.category = categorySelect.value;
          }
          row.categoryIsNew = !!row.category && !(fileData.categories || []).includes(row.category);
          renderImportStatusCell(statusCell, row);
        };
        categorySelect.addEventListener('change', syncCategory);
        newCategoryInput.addEventListener('input', syncCategory);
        categoryField.appendChild(categorySelect);
        categoryField.appendChild(newCategoryInput);
        fields.appendChild(categoryField);

        const splitField = document.createElement('div');
        splitField.className = 'col';
        const splitLabel = document.createElement('label');
        splitLabel.textContent = 'Split';
        splitField.appendChild(splitLabel);

        const splitSelect = document.createElement('select');
        [
          {value:'equal', label:'Equal split'},
          {value:'payer', label:'Payer only'}
        ].forEach(optionConfig => {
          const option = document.createElement('option');
          option.value = optionConfig.value;
          option.textContent = optionConfig.label;
          splitSelect.appendChild(option);
        });
        splitSelect.value = row.splitChoice;
        splitSelect.addEventListener('change', () => {
          row.splitChoice = splitSelect.value;
          renderImportStatusCell(statusCell, row);
        });
        splitField.appendChild(splitSelect);
        fields.appendChild(splitField);

        card.appendChild(fields);

        const statusCell = document.createElement('div');
        statusCell.className = 'import-card-status';
        renderImportStatusCell(statusCell, row);
        card.appendChild(statusCell);

        list.appendChild(card);
      });
      importPreview.appendChild(list);

      const actions = document.createElement('div');
      actions.className = 'import-actions';
      const importBtn = document.createElement('button');
      importBtn.className = 'btn';
      importBtn.type = 'button';
      importBtn.textContent = 'Import selected';
      importBtn.addEventListener('click', importSelectedRows);
      const selectAllBtn = document.createElement('button');
      selectAllBtn.className = 'btn ghost';
      selectAllBtn.type = 'button';
      selectAllBtn.textContent = 'Select all shown';
      selectAllBtn.addEventListener('click', () => {
        importDraftRows.forEach(row => { if(row.month === selectedMonth) row.selected = true; });
        renderImportPreview();
      });
      actions.appendChild(importBtn);
      actions.appendChild(selectAllBtn);
      importPreview.appendChild(actions);
    }
    async function importSelectedRows(){
      if(importInProgress) return;
      const selectedMonth = normalizeMonthKey(monthEl.value);
      const rows = importDraftRows.filter(row => row.selected && row.month === selectedMonth && row.amount > 0);
      if(!rows.length){
        alert('No valid rows are selected for this month.');
        return;
      }
      const newCategories = uniqueList(rows.map(row => row.category).filter(category => category && !(fileData.categories || []).includes(category)));
      if(newCategories.length){
        const message = `These new categories will be added:\n\n${newCategories.join('\n')}\n\nContinue?`;
        if(!confirm(message)) return;
        fileData.categories = uniqueList([...(fileData.categories || []), ...newCategories]);
        saveGlobalToLocal();
      }
      const missingPayer = rows.find(row => !row.payer);
      if(missingPayer){
        alert('Every imported row needs a payer. Please choose a payer in the preview.');
        return;
      }
      const missingCategory = rows.find(row => !row.category);
      if(missingCategory){
        alert('Every imported row needs a category. Please choose a category in the preview.');
        return;
      }
      const duplicateRows = rows.filter(row => findImportDuplicate(row) >= 0);
      if(duplicateRows.length){
        const proceed = confirm(`${duplicateRows.length} selected row(s) look like duplicates already in this month. Import them anyway?`);
        if(!proceed) return;
      }
      importInProgress = true;
      try{
        for(const row of rows){
          (data.expenses = data.expenses || []).push(importRowToExpense(row));
        }
        await save();
        importDraftRows = importDraftRows.filter(row => !rows.includes(row));
        renderAll();
        renderImportPreview();
        alert(`Imported ${rows.length} ${rows.length === 1 ? 'expense' : 'expenses'}.`);
      } finally {
        importInProgress = false;
      }
    }
    function clearImport(){
      importRowsRaw = [];
      importDraftRows = [];
      if(importFile) importFile.value = '';
      if(importMapping) importMapping.style.display = 'none';
      if(importPreview){ importPreview.style.display = 'none'; importPreview.innerHTML = ''; }
    }

    async function addQuickExpenses(){
      const payer = quickPerson.value;
      const category = quickCategory.value;
      if(!payer || !category){ alert('Please select a person and category'); return; }
      const amounts = parseAmounts(quickAmounts.value);
      if(!amounts.length){ alert('Add at least one amount'); return; }
      const note = document.getElementById('quickNote') ? document.getElementById('quickNote').value.trim() : '';
      const splits = getSplitsFromSliders(quickSplitSliders);
      const date = quickDate?.value || dateValueFromMonth(monthEl.value);
      for(const a of amounts){ if(a<=0) continue; (data.expenses = data.expenses||[]).push({date, payer, category, amount: round2(+a), note, splits}); }
      await save(); renderAll(); quickAmounts.value='';
    }

    function makeBatchRow(category=''){
      const row=document.createElement('div'); row.className='row';
      row.innerHTML='\
        <div class="col" style="flex:1;min-width:160px"><label>Category</label><select class="batchCategory"></select></div>\
        <div class="col" style="width:160px"><label>Amount</label><input class="batchAmount" type="number" inputmode="decimal" step="0.01" placeholder="0.00"/></div>\
        <div class="col" style="flex:1;min-width:200px"><label>Note (optional)</label><input class="batchNote" type="text" placeholder="e.g., receipt #1234"/></div>\
        <div class="expense-added" aria-live="polite" aria-atomic="true"></div>';
      const sel=row.querySelector('.batchCategory'); fillSelect(sel, (fileData.categories||[])); if(category) sel.value=category;
      return row;
    }
    async function saveBatchExpense(){
      const payer=batchPerson.value; if(!payer){ alert('Select a person'); return; }
      let row=batchRows.querySelector('.row');
      if(!row){
        row = makeBatchRow();
        batchRows.appendChild(row);
      }
      const splits = getSplitsFromSliders(batchSplitSliders);
      const categorySelect = row.querySelector('.batchCategory');
      const cat=categorySelect.value;
      const amountInput = row.querySelector('.batchAmount');
      const noteInput = row.querySelector('.batchNote');
      const addedStatus = row.querySelector('.expense-added');
      const val=parseFloat(amountInput.value);
      const note=(noteInput?.value||'').trim();
      if(!cat){ alert('Please select a category'); return; }
      if(!isFinite(val) || val<=0){ alert('Please enter a valid amount'); return; }
      const date = batchDate?.value || dateValueFromMonth(monthEl.value);
      (data.expenses = data.expenses||[]).push({date, payer, category:cat, amount: round2(val), note, splits});
      await save();
      renderAll();
      categorySelect.value='';
      amountInput.value='';
      if(noteInput) noteInput.value='';
      if(addedStatus){
        addedStatus.textContent='✓';
        addedStatus.classList.add('show');
        setTimeout(()=> addedStatus.classList.remove('show'), 1200);
        setTimeout(()=> { addedStatus.textContent=''; }, 1500);
      }
      amountInput.focus();
    }

    // ---------- Init & events ----------
    document.addEventListener('DOMContentLoaded', async ()=>{
      initSupabase();
      if(supabaseClient){
        await refreshSession();

        // Handle email confirmation and password reset links
        const handleAuthCallback = async () => {
          const hashParams = new URLSearchParams(window.location.hash.substring(1));
          const type = hashParams.get('type');
          const error = hashParams.get('error');
          const errorDescription = hashParams.get('error_description');

          if(error){
            showAuthMessage(decodeURIComponent(errorDescription || error), 'error');
            // Clean up URL
            window.history.replaceState({}, document.title, window.location.pathname);
            return;
          }

          if(type === 'signup'){
            showAuthMessage('Email confirmed! You can now sign in.', 'success');
            window.history.replaceState({}, document.title, window.location.pathname);
          } else if(type === 'recovery'){
            const newPassword = prompt('Enter your new password (min. 6 characters):');
            if(newPassword && validatePassword(newPassword)){
              try {
                const { error: updateError } = await supabaseClient.auth.updateUser({ password: newPassword });
                if(updateError){
                  showAuthMessage(updateError.message, 'error');
                } else {
                  showAuthMessage('Password updated successfully! You are now signed in.', 'success');
                  authPassword.value = '';
                }
              } catch(err){
                showAuthMessage('Failed to update password. Please try again.', 'error');
              }
            } else if(newPassword){
              showAuthMessage('Password must be at least 6 characters long.', 'error');
            }
            window.history.replaceState({}, document.title, window.location.pathname);
          }
        };

        // Check for auth callback on page load
        if(window.location.hash){
          await handleAuthCallback();
        }

        supabaseClient.auth.onAuthStateChange(async (event, session) => {
          currentSession = session || null;
          currentUser = currentSession?.user || null;
          if(currentUser){
            const loaded = await loadRemote();
            if(loaded){
              await loadFromFileOrLocal(monthEl.value);
            }
          }
          updateAuthStatus();
          renderAll();
          renderAverages();
          refreshAppView();

          // Handle specific auth events
          if(event === 'SIGNED_IN' && !window.location.hash){
            // Only show if not from callback (callback shows its own message)
            authStatus.textContent = `Signed in as ${currentUser?.email || ''}`;
          } else if(event === 'SIGNED_OUT'){
            authStatus.textContent = 'Not signed in.';
          }
        });
      }
      const initial = thisMonthValue();
      monthEl.value = initial;
      syncEntryDatesToMonth();
      if(currentUser){
        const loaded = await loadRemote();
        if(loaded){
          await loadFromFileOrLocal(initial);
        }
      } else {
        // Load global people/categories from local storage if not using Supabase
        loadGlobalFromLocal();
        await loadFromFileOrLocal(initial);
      }
      lastLoadedMonth = initial;
      loadTheme();
      setupCollapsibles();
      renderAll();
      renderAverages();
      updateAuthStatus();
      updateCloudStatus();
      if(currentUser) activeView = 'home';
      refreshAppView();
      if(!batchRows.children.length) batchRows.appendChild(makeBatchRow());

      if(analysisView) analysisView.addEventListener('change', renderAverages);
      if(analysisCompareA) analysisCompareA.addEventListener('change', renderAverages);
      if(analysisCompareB) analysisCompareB.addEventListener('change', renderAverages);
      if(menuToggle){
        menuToggle.addEventListener('click', toggleMobileMenu);
      }
      if(mobileMenuSignInBtn){
        mobileMenuSignInBtn.addEventListener('click', () => {
          setView('account');
          closeMobileMenu();
          authEmail?.focus();
        });
      }
      if(mobileMenuSignOutBtn){
        mobileMenuSignOutBtn.addEventListener('click', () => {
          closeMobileMenu();
          signOutBtn?.click();
        });
      }
      document.querySelectorAll('[data-target-view]').forEach(button => {
        button.addEventListener('click', () => {
          setView(button.dataset.targetView);
          closeMobileMenu();
        });
      });

      monthEl.addEventListener('change', async ()=>{
        const target = monthEl.value;
        await switchMonthSafely(target);
        syncEntryDatesToMonth();
        buildSplitSliders(quickSplitSliders);
        buildSplitSliders(batchSplitSliders);
      });

      addPersonBtn.addEventListener('click', async ()=>{
        const v = personNameEl.value; personNameEl.value=''; await addPerson(v); personNameEl.focus();
        buildSplitSliders(quickSplitSliders); buildSplitSliders(batchSplitSliders);
      });
      personNameEl.addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ addPersonBtn.click(); }});

      addCategoryBtn.addEventListener('click', async ()=>{
        const v = categoryNameEl.value; categoryNameEl.value=''; await addCategory(v); categoryNameEl.focus();
      });

      modeQuickBtn.addEventListener('click', ()=>{
        modeQuickBtn.classList.add('active'); modeBatchBtn.classList.remove('active'); modeImportBtn.classList.remove('active');
        quickPanel.style.display='block'; batchPanel.style.display='none'; importPanel.style.display='none';
      });
      modeBatchBtn.addEventListener('click', ()=>{
        modeBatchBtn.classList.add('active'); modeQuickBtn.classList.remove('active'); modeImportBtn.classList.remove('active');
        quickPanel.style.display='none'; batchPanel.style.display='block'; importPanel.style.display='none';
        if(!batchRows.children.length) batchRows.appendChild(makeBatchRow());
      });
      modeImportBtn.addEventListener('click', ()=>{
        modeImportBtn.classList.add('active'); modeBatchBtn.classList.remove('active'); modeQuickBtn.classList.remove('active');
        quickPanel.style.display='none'; batchPanel.style.display='none'; importPanel.style.display='block';
      });

      addQuickBtn.addEventListener('click', addQuickExpenses);
      clearQuickBtn.addEventListener('click', ()=>{ quickAmounts.value=''; });

      saveBatchBtn.addEventListener('click', saveBatchExpense);

      clearMonthBtn.addEventListener('click', async ()=>{
        const msg = "Are you sure you want to delete this month's expenses? You won't be able to get them back.";
        if(confirm(msg)){
          const secondMsg = `This will also delete ${monthLabel(monthEl.value)} from the cloud if you are signed in. Type DELETE to confirm.`;
          if(currentUser && prompt(secondMsg) !== 'DELETE') return;
          data.expenses = [];
          await save({ allowEmptyOverwrite: true });
          renderAll();
        }
      });
      if(checkDuplicatesBtn){
        checkDuplicatesBtn.addEventListener('click', checkDuplicateExpenses);
      }
      if(importFile){
        importFile.addEventListener('change', handleImportFileChange);
      }
      if(buildImportPreviewBtn){
        buildImportPreviewBtn.addEventListener('click', buildImportDraftRows);
      }
      if(clearImportBtn){
        clearImportBtn.addEventListener('click', clearImport);
      }
      if(importUseSamePayer){
        importUseSamePayer.addEventListener('change', () => {
          if(importDraftRows.length) buildImportDraftRows();
        });
      }
      if(importSamePayer){
        importSamePayer.addEventListener('change', () => {
          if(importUseSamePayer?.checked && importDraftRows.length) buildImportDraftRows();
        });
      }
      if(saveCloudBtn){
        saveCloudBtn.addEventListener('click', () => syncCloudNow());
      }
      if(retryCloudBtn){
        retryCloudBtn.addEventListener('click', retryCloudSync);
      }
      if(exportDataBtn){
        exportDataBtn.addEventListener('click', exportAllData);
      }
      if(loadBackupsBtn){
        loadBackupsBtn.addEventListener('click', loadAndRenderBackups);
      }
      if(migrateDatesBtn){
        migrateDatesBtn.addEventListener('click', migrateMissingExpenseDates);
      }
      if(closeExpenseDetailBtn){
        closeExpenseDetailBtn.addEventListener('click', closeExpenseDetail);
      }
      if(editExpenseDetailBtn){
        editExpenseDetailBtn.addEventListener('click', editExpenseFromDetail);
      }
      if(saveExpenseDetailBtn){
        saveExpenseDetailBtn.addEventListener('click', saveExpenseDetailEdits);
      }
      if(cancelExpenseEditBtn){
        cancelExpenseEditBtn.addEventListener('click', cancelExpenseEdit);
      }
      if(deleteExpenseDetailBtn){
        deleteExpenseDetailBtn.addEventListener('click', deleteExpenseFromDetail);
      }
      if(expenseDetailModal){
        expenseDetailModal.addEventListener('click', event => {
          if(event.target === expenseDetailModal) closeExpenseDetail();
        });
      }
      document.addEventListener('keydown', event => {
        if(event.key === 'Escape' && expenseDetailModal?.style.display !== 'none'){
          closeExpenseDetail();
        }
      });

      signInBtn.addEventListener('click', async ()=>{
        hideAuthMessage();
        if(!supabaseClient){
          showAuthMessage('Supabase is not configured. Please add your credentials to the code.', 'error');
          return;
        }
        const email = authEmail.value.trim();
        const password = authPassword.value;

        // Validation
        if(!email || !password){
          showAuthMessage('Please enter both email and password.', 'error');
          return;
        }
        if(!validateEmail(email)){
          showAuthMessage('Please enter a valid email address.', 'error');
          return;
        }

        setButtonLoading(signInBtn, true);
        try {
          const { data: signInData, error } = await supabaseClient.auth.signInWithPassword({ email, password });
          if(error){
            showAuthMessage(error.message, 'error');
            return;
          }
          if(signInData?.session){
            currentSession = signInData.session;
            currentUser = signInData.session.user;
          }
          await refreshSession();
          const loaded = await loadRemote();
          if(!loaded){
            showAuthMessage('Signed in, but cloud data could not be loaded. Uploads are blocked so existing cloud data is protected.', 'error');
            return;
          }
          await loadFromFileOrLocal(monthEl.value);
          updateAuthStatus();
          renderAll();
          renderAverages();
          setView('home');
          showAuthMessage(`Successfully signed in as ${currentUser?.email || ''}`, 'success');
          authPassword.value = ''; // Clear password field
        } finally {
          setButtonLoading(signInBtn, false);
        }
      });

      signUpBtn.addEventListener('click', async ()=>{
        hideAuthMessage();
        if(!supabaseClient){
          showAuthMessage('Supabase is not configured. Please add your credentials to the code.', 'error');
          return;
        }
        const email = authEmail.value.trim();
        const password = authPassword.value;

        // Validation
        if(!email || !password){
          showAuthMessage('Please enter both email and password.', 'error');
          return;
        }
        if(!validateEmail(email)){
          showAuthMessage('Please enter a valid email address.', 'error');
          return;
        }
        if(!validatePassword(password)){
          showAuthMessage('Password must be at least 6 characters long.', 'error');
          return;
        }

        setButtonLoading(signUpBtn, true);
        try {
          const { data, error } = await supabaseClient.auth.signUp({ email, password });
          if(error){
            showAuthMessage(error.message, 'error');
            return;
          }

          // Check if email confirmation is required
          if(data?.user && !data.session){
            showAuthMessage('Account created! Check your email to confirm your account, then sign in.', 'success');
            authPassword.value = ''; // Clear password field
          } else if(data?.session){
            // Auto-confirmation is enabled, user is already signed in
            showAuthMessage('Account created and signed in successfully!', 'success');
            authPassword.value = '';
          }
        } finally {
          setButtonLoading(signUpBtn, false);
        }
      });

      signOutBtn.addEventListener('click', async ()=>{
        if(!supabaseClient){ return; }
        hideAuthMessage();
        if(currentUser && cloudSavePending){
          const shouldSync = confirm('You have local changes that may not be saved to cloud yet. Save to cloud before signing out?');
          if(shouldSync){
            const synced = await syncCloudNow();
            if(!synced) return;
          }
        }
        setButtonLoading(signOutBtn, true);
        let signOutError = null;
        try {
          await withTimeout(supabaseClient.auth.signOut(), 5000, 'Sign out');
        } catch(err){
          signOutError = err;
          console.warn('Sign out failed', err);
        }
        currentUser = null;
        currentSession = null;
        fileData = { people: [], categories: [], months: {} };
        // Load global data from local storage after signing out
        loadGlobalFromLocal();
        updateAuthStatus();
        renderAll();
        renderAverages();
        refreshAppView();
        if(signOutError){
          showAuthMessage('Signed out locally. Cloud sign-out may have timed out.', 'info');
        } else {
          showAuthMessage('Successfully signed out.', 'success');
        }
        authPassword.value = ''; // Clear password field
        setButtonLoading(signOutBtn, false);
      });

      forgotPasswordBtn.addEventListener('click', async ()=>{
        hideAuthMessage();
        if(!supabaseClient){
          showAuthMessage('Supabase is not configured. Please add your credentials to the code.', 'error');
          return;
        }
        const email = authEmail.value.trim();

        // Validation
        if(!email){
          showAuthMessage('Please enter your email address.', 'error');
          return;
        }
        if(!validateEmail(email)){
          showAuthMessage('Please enter a valid email address.', 'error');
          return;
        }

        if(!confirm(`Send password reset email to ${email}?`)){
          return;
        }

        setButtonLoading(forgotPasswordBtn, true);
        try {
          const { error } = await supabaseClient.auth.resetPasswordForEmail(email, {
            redirectTo: window.location.origin + window.location.pathname
          });
          if(error){
            showAuthMessage(error.message, 'error');
            return;
          }
          showAuthMessage('Password reset email sent! Check your inbox.', 'success');
        } finally {
          setButtonLoading(forgotPasswordBtn, false);
        }
      });

      // Keyboard support: Enter key on password field triggers sign in
      authPassword.addEventListener('keydown', (e)=>{
        if(e.key === 'Enter'){
          e.preventDefault();
          signInBtn.click();
        }
      });

      // Keyboard support: Enter key on email field focuses password
      authEmail.addEventListener('keydown', (e)=>{
        if(e.key === 'Enter'){
          e.preventDefault();
          authPassword.focus();
        }
      });

      if(themeToggle){
        themeToggle.addEventListener('change', (e)=>{
          saveTheme(e.target.checked);
        });
      }

      // initial split options
      buildSplitSliders(quickSplitSliders);
      buildSplitSliders(batchSplitSliders);
    });
