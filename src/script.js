        document.addEventListener('DOMContentLoaded', function () {
            // --- DOM Elements ---
            const appContainer = document.getElementById('app-container');
            const editModeToggle = document.getElementById('edit-mode-toggle');
            const summaryButton = document.getElementById('summary-button');
            const resetDataButton = document.getElementById('reset-data-button');
            const searchBox = document.getElementById('search-box');
            const exportJsonButton = document.getElementById('export-json-button');
            const exportExcelButton = document.getElementById('export-excel-button');
            const exportWordButton = document.getElementById('export-word-button');
            if (exportJsonButton) {
                exportJsonButton.textContent = 'エクスポート (JSON)';
            }
            if (exportExcelButton) {
                exportExcelButton.textContent = 'エクスポート (Excel)';
            }
            if (exportWordButton) {
                exportWordButton.textContent = 'エクスポート (Word)';
            }
            const importJsonButton = document.getElementById('import-json-button');
            const bulkEditButton = document.getElementById('bulk-edit-button');
            const bulkAddButton = document.getElementById('bulk-add-button');

            const bulkAddModal = document.getElementById('bulk-add-modal');
            const saveBulkAddButton = document.getElementById('save-bulk-add-button');
            const cancelBulkAddButton = document.getElementById('cancel-bulk-add-button');
            const bulkAddCloseButton = bulkAddModal.querySelector('.close-btn');

            const bulkAddItemSelect = document.getElementById('bulk-add-item-select');
            const bulkAddMonthsContainer = document.getElementById('bulk-add-months-container');
            const bulkAddTypeIncome = document.getElementById('bulk-add-type-income');
            const bulkAddTypeExpenditure = document.getElementById('bulk-add-type-expenditure');
            const bulkAddSelectAllMonths = document.getElementById('bulk-add-select-all-months');
            const bulkAddSelectAllLabel = document.querySelector('label[for="bulk-add-select-all-months"] strong');
            if (bulkAddSelectAllLabel) {
                bulkAddSelectAllLabel.textContent = 'すべて選択';
            }

            let bulkAddCandidates = [];
            const bulkAddCurrencyFormatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });

            const getSelectedBulkAddItem = () => {
                if (!bulkAddItemSelect) return null;
                return bulkAddCandidates.find(candidate => candidate.value === bulkAddItemSelect.value) || null;
            };

            const populateBulkAddItemOptions = () => {
                if (!bulkAddItemSelect) return;
                const type = bulkAddTypeIncome.checked ? 'income' : 'expenditure';
                bulkAddItemSelect.innerHTML = '';

                const placeholderOption = document.createElement('option');
                placeholderOption.value = '';
                placeholderOption.textContent = '追加する項目を選択';
                placeholderOption.disabled = true;
                placeholderOption.selected = true;
                bulkAddItemSelect.appendChild(placeholderOption);

                const filteredCandidates = bulkAddCandidates.filter(candidate => candidate.type === type);
                const hasCandidates = filteredCandidates.length > 0;
                bulkAddItemSelect.disabled = !hasCandidates;

                if (filteredCandidates.length === 0) {
                    const emptyOption = document.createElement('option');
                    emptyOption.value = '';
                    emptyOption.disabled = true;
                    emptyOption.textContent = type === 'income' ? '追加できる収入項目がありません' : '追加できる支出項目がありません';
                    bulkAddItemSelect.appendChild(emptyOption);
                    return;
                }

                filteredCandidates.forEach(candidate => {
                    const option = document.createElement('option');
                    option.value = candidate.value;
                    option.textContent = `${candidate.yearLabel} ${candidate.monthLabel} - ${candidate.name} (${bulkAddCurrencyFormatter.format(candidate.amount)})`;
                    bulkAddItemSelect.appendChild(option);
                });
            };

            const updateMonthCheckboxes = () => {
                if (!bulkAddMonthsContainer) return;
                const type = bulkAddTypeIncome.checked ? 'income' : 'expenditure';
                const selectedItem = getSelectedBulkAddItem();

                let allEnabledAreChecked = true;
                let hasEnabledCheckboxes = false;

                const checkboxes = bulkAddMonthsContainer.querySelectorAll('input[type="checkbox"]');
                checkboxes.forEach(checkbox => {
                    const label = checkbox.nextElementSibling;

                    if (!selectedItem || selectedItem.type !== type) {
                        checkbox.checked = false;
                        checkbox.disabled = true;
                        if (label) label.classList.add('disabled');
                        return;
                    }

                    const [yearId, monthId] = checkbox.value.split('|');
                    const monthData = findMonth(yearId, monthId);
                    const itemExists = !monthData || monthData[type].some(item => item.name === selectedItem.name);

                    checkbox.disabled = itemExists;
                    if (itemExists) {
                        checkbox.checked = false;
                        if (label) label.classList.add('disabled');
                    } else {
                        hasEnabledCheckboxes = true;
                        checkbox.checked = true;
                        if (label) label.classList.remove('disabled');
                    }

                    if (!checkbox.disabled && !checkbox.checked) {
                        allEnabledAreChecked = false;
                    }
                });

                bulkAddSelectAllMonths.disabled = !hasEnabledCheckboxes;
                if (!hasEnabledCheckboxes) {
                    bulkAddSelectAllMonths.checked = false;
                } else {
                    bulkAddSelectAllMonths.checked = hasEnabledCheckboxes && allEnabledAreChecked;
                }
                saveBulkAddButton.disabled = !selectedItem || !hasEnabledCheckboxes;
            };

            const buildBulkAddCandidates = () => {
                const byKey = new Map(); // key: `${type}|${name}`
                state.financialData.forEach(yearData => {
                    yearData.months.forEach(monthData => {
                        ['income', 'expenditure'].forEach(type => {
                            monthData[type].forEach(item => {
                                const key = `${type}|${item.name}`;
                                if (byKey.has(key)) return;
                                byKey.set(key, {
                                    value: key,
                                    type,
                                    name: item.name,
                                    amount: Number(item.amount) || 0,
                                    yearLabel: yearData.year,
                                    monthLabel: monthData.month,
                                    data: { ...item }
                                });
                            });
                        });
                    });
                });
                return Array.from(byKey.values()).sort((a, b) => a.name.localeCompare(b.name, 'ja'));
            };

            const populateBulkAddModal = () => {
                if (!bulkAddMonthsContainer) return;
                bulkAddMonthsContainer.innerHTML = '';
                state.financialData.forEach(yearData => {
                    const yearGroup = document.createElement('div');
                    yearGroup.className = 'month-checkbox-group';

                    const yearTitle = document.createElement('strong');
                    yearTitle.className = 'month-checkbox-group-title';
                    yearTitle.textContent = yearData.year;
                    yearGroup.appendChild(yearTitle);

                    yearData.months.forEach(monthData => {
                        const monthId = `${yearData.id}|${monthData.id}`;
                        const item = document.createElement('div');
                        item.className = 'month-checkbox-item';

                        const checkbox = document.createElement('input');
                        checkbox.type = 'checkbox';
                        checkbox.id = `month-checkbox-${monthId}`;
                        checkbox.value = monthId;

                        const label = document.createElement('label');
                        label.htmlFor = `month-checkbox-${monthId}`;
                        label.textContent = `${monthData.month}月`;

                        item.appendChild(checkbox);
                        item.appendChild(label);
                        yearGroup.appendChild(item);
                    });
                    bulkAddMonthsContainer.appendChild(yearGroup);
                });

                bulkAddCandidates = buildBulkAddCandidates();
                populateBulkAddItemOptions();
                updateMonthCheckboxes();
            };

            bulkAddTypeIncome.addEventListener('change', () => {
                populateBulkAddItemOptions();
                updateMonthCheckboxes();
            });
            bulkAddTypeExpenditure.addEventListener('change', () => {
                populateBulkAddItemOptions();
                updateMonthCheckboxes();
            });
            if (bulkAddItemSelect) {
                bulkAddItemSelect.addEventListener('change', updateMonthCheckboxes);
            }

            bulkAddSelectAllMonths.addEventListener('change', (e) => {
                const isChecked = e.target.checked;
                const checkboxes = bulkAddMonthsContainer.querySelectorAll('input[type="checkbox"]');
                checkboxes.forEach(checkbox => {
                    if (!checkbox.disabled) {
                        checkbox.checked = isChecked;
                    }
                });
            });

            bulkAddButton.addEventListener('click', () => {
                populateBulkAddModal();
                showModal(bulkAddModal);
            });

            cancelBulkAddButton.addEventListener('click', () => hideModal(bulkAddModal));
            bulkAddCloseButton.addEventListener('click', () => hideModal(bulkAddModal));

            saveBulkAddButton.addEventListener('click', () => {
                const selectedItem = getSelectedBulkAddItem();
                if (!selectedItem) {
                    showPopupMessage('追加する項目を選択してください。');
                    return;
                }

                const selectedMonthsCheckboxes = bulkAddMonthsContainer.querySelectorAll('input[type="checkbox"]:checked:not(:disabled)');

                if (selectedMonthsCheckboxes.length === 0) {
                    showPopupMessage('最低でも1つの対象月を選択してください。');
                    return;
                }

                let addedCount = 0;
                selectedMonthsCheckboxes.forEach(checkbox => {
                    const [yearId, monthId] = checkbox.value.split('|');
                    const monthData = findMonth(yearId, monthId);
                    if (!monthData) return;

                    const duplicate = monthData[selectedItem.type].some(item => item.name === selectedItem.name);
                    if (duplicate) return;

                    const newItem = {
                        ...selectedItem.data,
                        id: `item_${Date.now()}_${Math.random()}`
                    };
                    newItem.amount = Number(newItem.amount) || 0;
                    monthData[selectedItem.type].push(newItem);
                    addedCount++;
                });

                if (addedCount > 0) {
                    rerender();
                    showPopupMessage(`${addedCount}件の月に「${selectedItem.name}」を追加しました。`);
                    hideModal(bulkAddModal);
                } else {
                    showPopupMessage('追加できる月がありませんでした。');
                }
            });
            const settingsButton = document.getElementById('settings-button');

            // Smartphone Warning Modal
            const smartphoneWarningModal = document.getElementById('smartphone-warning-modal');
            const smartphoneWarningCloseButton = document.getElementById('smartphone-warning-close-button');
            const smartphoneWarningHeaderCloseButton = smartphoneWarningModal.querySelector('.close-btn');

            const closeSmartphoneWarning = () => {
                smartphoneWarningModal.style.display = 'none';
                localStorage.setItem('smartphoneWarningShown', 'true');
            };

            smartphoneWarningCloseButton.addEventListener('click', closeSmartphoneWarning);
            if (smartphoneWarningHeaderCloseButton) {
                smartphoneWarningHeaderCloseButton.addEventListener('click', closeSmartphoneWarning);
            }


            const userNameDisplay = document.getElementById('user-name-display');
            const userNameEdit = document.getElementById('user-name-edit');

            // Password Modal
            const passwordModal = document.getElementById('password-modal');
            const passwordForm = document.getElementById('password-form');
            const passwordInput = document.getElementById('password-input');
            const passwordError = document.getElementById('password-error');
            const createNewDataLink = document.getElementById('create-new-data-from-lock-screen');
            const rememberPasswordCheckbox = document.getElementById('remember-password');
            const toggleMainPasswordVisibility = document.getElementById('toggle-main-password-visibility');
            // Remember password is intentionally disabled for security; hide its UI if present.
            if (rememberPasswordCheckbox && rememberPasswordCheckbox.closest('.form-group')) {
                rememberPasswordCheckbox.closest('.form-group').style.display = 'none';
            }

            createNewDataLink.addEventListener('click', (e) => {
                e.preventDefault();
                if (confirm('パスワードが不明なため、新しいデータを作成します。よろしいですか？\n注意：現在のデータは失われます。')) {
                    // Clear all stored data to ensure a fresh start
                    localStorage.removeItem('financialData');
                    localStorage.removeItem('userName');
                    localStorage.removeItem('password');
                    localStorage.removeItem('categories');
                    localStorage.removeItem('rememberedPassword'); // 記憶されたパスワードも削除
                    localStorage.removeItem('dataIsEncrypted');

                    // Reset the state in the app
                    state.financialData = [];
                    state.userName = '';
                    state.password = '';
                    state.categories = loadCategories(); // Load default/empty categories

                    // Hide the password modal
                    passwordModal.classList.remove('visible');

                    // Show the new data creation modal
                    showModal(newDataModal);
                }
            });

            // Edit Modal
            const modal = document.getElementById('modal');
            const modalForm = document.getElementById('modal-form');
            const modalTitle = document.getElementById('modal-title');
            const itemNameInput = document.getElementById('item-name');
            const itemAmountInput = document.getElementById('item-amount');
            const cancelButton = document.getElementById('cancel-button');
            const modalCloseButton = modal.querySelector('.close-btn');
            if (modalCloseButton) {
                modalCloseButton.addEventListener('click', () => hideModal(modal));
            }

            // Bulk Edit Modal
            const bulkEditModal = document.getElementById('bulk-edit-modal');
            const bulkEditForm = document.getElementById('bulk-edit-form');
            const bulkEditSelect = document.getElementById('bulk-edit-select');
            const bulkEditNewName = document.getElementById('bulk-edit-new-name');
            const bulkEditNewAmount = document.getElementById('bulk-edit-new-amount');
            const bulkCancelButton = document.getElementById('bulk-cancel-button');
            // Bulk Add Modal (追加)
            const bulkAddContent = document.getElementById('bulk-add-content');

            // Summary Modal
            const summaryModal = document.getElementById('summary-modal');
            const summaryCloseButton = document.getElementById('summary-close-button');
            const summaryContent = document.getElementById('summary-content');

            // Category Summary Modal
            const categorySummaryModal = document.getElementById('category-summary-modal');
            const categorySummaryButton = document.getElementById('category-summary-button');
            const categorySummaryCloseButton = document.getElementById('category-summary-close-button');
            const categorySummaryContent = document.getElementById('category-summary-content');

            // New Data Modal
            const newDataButton = document.getElementById('new-data-button');
            const newDataModal = document.getElementById('new-data-modal');
            const newDataForm = document.getElementById('new-data-form');
            const newDataCancelButton = document.getElementById('new-data-cancel-button');
            const newDataCloseButton = document.getElementById('new-data-close-button');
            const startYearInput = document.getElementById('start-year');
            const initialBalanceInput = document.getElementById('initial-balance');
            const userNameInput = document.getElementById('user-name');

            // Add Year Modal
            const addYearModal = document.getElementById('add-year-modal');
            const addYearForm = document.getElementById('add-year-form');
            const addYearCloseButton = document.getElementById('add-year-close-button');
            const addYearCancelButton = document.getElementById('add-year-cancel-button');
            const newYearNameInput = document.getElementById('new-year-name');
            const addYearMonthsContainer = document.getElementById('add-year-months-container');

            // Edit Months Modal
            const editMonthsModal = document.getElementById('edit-months-modal');
            const editMonthsForm = document.getElementById('edit-months-form');
            const editMonthsTitle = document.getElementById('edit-months-title');
            const editMonthsContainer = document.getElementById('edit-months-container');
            const editMonthsCloseButton = document.getElementById('edit-months-close-button');
            const editMonthsCancelButton = document.getElementById('edit-months-cancel-button');
            const deleteYearButton = document.getElementById('delete-year-button');

            // Category Settings Modal (now part of general settings modal)
            const settingsModal = document.getElementById('settings-modal');
            const settingsCloseButton = document.getElementById('settings-close-button');
            const categoryInput = document.getElementById('category-input');
            const saveSettingsButton = document.getElementById('save-settings');
            const cancelCategoriesButton = document.getElementById('cancel-categories-button');
            const autoGetCategoriesButton = document.getElementById('auto-get-categories-button');
            const passwordSettingInput = document.getElementById('password-setting');
            const requirePasswordToggle = document.getElementById('require-password-on-load');
            const requirePasswordToggleGroup = document.getElementById('require-password-toggle-group');
            const togglePasswordVisibility = document.getElementById('toggle-password-visibility');

            // --- App State ---
            let state = {
                isEditMode: false,
                financialData: [],
                userName: '',
                editingItem: null, // { yearId, monthId, type, itemId }
                searchTerm: '',
                categories: [],
                password: '',
                requirePasswordOnLoad: false
            };

            // --- SVG Icons ---
            const icons = {
                add: '<svg viewBox="0 0 24 24"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>',
                edit: '<svg viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>',
                delete: '<svg viewBox="0 0 24 24"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>',
            };

            // --- Event Listeners ---
            if (toggleMainPasswordVisibility) {
                toggleMainPasswordVisibility.addEventListener('click', function () {
                    const type = passwordInput.getAttribute('type') === 'password' ? 'text' : 'password';
                    passwordInput.setAttribute('type', type);

                    const eyeOpen = this.querySelector('.eye-icon-open');
                    const eyeClosed = this.querySelector('.eye-icon-closed');

                    if (type === 'password') {
                        eyeOpen.style.display = 'block';
                        eyeClosed.style.display = 'none';
                    } else {
                        eyeOpen.style.display = 'none';
                        eyeClosed.style.display = 'block';
                    }
                });
            }
            if (togglePasswordVisibility) {
                togglePasswordVisibility.addEventListener('click', function () {
                    const type = passwordSettingInput.getAttribute('type') === 'password' ? 'text' : 'password';
                    passwordSettingInput.setAttribute('type', type);

                    const eyeOpen = this.querySelector('.eye-icon-open');
                    const eyeClosed = this.querySelector('.eye-icon-closed');

                    if (type === 'text') {
                        eyeOpen.style.display = 'none';
                        eyeClosed.style.display = 'block';
                    } else {
                        eyeOpen.style.display = 'block';
                        eyeClosed.style.display = 'none';
                    }
                });
            }

            // --- Smartphone Warning ---
            function isSmartphone() {
                return /Android|webOS|iPhone|iPad|IEMobile|Opera Mini/i.test(navigator.userAgent);
            }

            function showSmartphoneWarning() {
                if (isSmartphone() && !localStorage.getItem('smartphoneWarningShown')) {
                    smartphoneWarningModal.style.display = 'flex';
                }
            }

            smartphoneWarningCloseButton.addEventListener('click', () => {
                smartphoneWarningModal.style.display = 'none';
                localStorage.setItem('smartphoneWarningShown', 'true');
            });

            // --- Data Initialization ---
            async function initializeApp() {
                showSmartphoneWarning(); // Check for smartphone on app init

                // Clear any previously persisted passwords for safety
                localStorage.removeItem('password');
                localStorage.removeItem('rememberedPassword');
                localStorage.removeItem('dataIsEncrypted');
                const storedRequirePassword = localStorage.getItem('requirePasswordOnLoad');
                const dataIsEncrypted = localStorage.getItem('dataIsEncrypted') === 'true';

                state.password = ''; // never persist password
                // If the value is stored as "false", it becomes false. Otherwise, it defaults to true.
                state.requirePasswordOnLoad = storedRequirePassword !== 'false';

                const hasStoredData = !!localStorage.getItem('financialData');
                if (dataIsEncrypted && hasStoredData) {
                    passwordModal.classList.add('visible');
                    passwordInput.value = '';
                    passwordInput.focus();
                } else {
                    // No password required, load data immediately
                    await loadAndRenderData();
                }
            }

            async function loadAndRenderData(enteredPassword = null) {
                const storedData = localStorage.getItem('financialData');
                const storedName = localStorage.getItem('userName');
                const dataIsEncrypted = localStorage.getItem('dataIsEncrypted') === 'true';

                if (storedName) {
                    state.userName = storedName;
                    userNameDisplay.textContent = state.userName;
                }

                state.categories = loadCategories();

                if (storedData) {
                    let parsedData = null;
                    let looksEncrypted = false;
                    try {
                        parsedData = JSON.parse(storedData);
                    } catch (err) {
                        looksEncrypted = true;
                    }

                    const passwordToUse = enteredPassword || state.password;
                    const mustRequestPassword = looksEncrypted || dataIsEncrypted;

                    if (mustRequestPassword && !passwordToUse) {
                        passwordModal.classList.add('visible');
                        passwordError.textContent = '';
                        passwordInput.focus();
                        return;
                    }

                    if (!parsedData && passwordToUse) {
                        try {
                            const bytes = CryptoJS.AES.decrypt(storedData, passwordToUse);
                            const decryptedString = bytes.toString(CryptoJS.enc.Utf8);
                            if (!decryptedString) {
                                throw new Error("Decryption failed. Invalid password?");
                            }
                            parsedData = JSON.parse(decryptedString);
                            state.password = passwordToUse; // keep only in memory
                        } catch (err) {
                            console.error('Failed to load data:', err);
                            passwordError.textContent = 'Password is invalid.';
                            passwordInput.value = '';
                            passwordInput.focus();
                            return;
                        }
                    }

                    if (!parsedData) {
                        console.error('Failed to load data: Parsed data is null');
                        if (confirm('Failed to load data. Reset to defaults?')) {
                            state.financialData = await loadDefaultData();
                        } else {
                            state.financialData = [];
                        }
                    } else {
                        state.financialData = processData(parsedData);
                    }
                } else {
                    state.financialData = await loadDefaultData();
                }

                // Hide the password modal
                if (passwordModal.classList.contains('visible')) {
                    passwordModal.style.opacity = '0';
                    setTimeout(() => {
                        passwordModal.classList.remove('visible');
                        passwordModal.style.opacity = '1';
                    }, 500);
                }
                rerender();
            }

            async function loadDefaultData() {
                try {
                    const response = await fetch('default-data.json');
                    if (!response.ok) {
                        throw new Error(`HTTP error! status: ${response.status}`);
                    }
                    const defaultData = await response.json();
                    return processData(defaultData);
                } catch (error) {
                    console.error('デフォルトデータの読み込みに失敗しました:', error);
                    showPopupMessage('ようこそ！');
                    return [];
                }
            }

            function processData(data) {
                return data.map(year => ({
                    ...year,
                    id: year.id || `year_${Date.now()}_${Math.random()}`,
                    months: year.months.map(month => ({
                        ...month,
                        id: month.id || `month_${Date.now()}_${Math.random()}`,
                        note: month.note || '',
                        startingBalance: Number(month.startingBalance) || 0,
                        income: (month.income || []).map(item => ({
                            ...item,
                            id: item.id || `item_${Date.now()}_${Math.random()}`,
                            amount: Number(item.amount) || 0
                        })),
                        expenditure: (month.expenditure || []).map(item => ({
                            ...item,
                            id: item.id || `item_${Date.now()}_${Math.random()}`,
                            amount: Number(item.amount) || 0
                        }))
                    }))
                }));
            }

            // --- State Management & Calculations ---
            const saveState = () => {
                let dataToStore = JSON.stringify(state.financialData);
                if (state.password) {
                    dataToStore = CryptoJS.AES.encrypt(dataToStore, state.password).toString();
                    state.requirePasswordOnLoad = true;
                }
                localStorage.setItem('financialData', dataToStore);
                localStorage.setItem('dataIsEncrypted', state.password ? 'true' : 'false');
                localStorage.setItem('userName', state.userName);
                localStorage.setItem('requirePasswordOnLoad', state.requirePasswordOnLoad);
                saveCategories(state.categories);
            };

            const saveCategories = (categories) => {
                localStorage.setItem('categories', JSON.stringify(categories));
            };

            const loadCategories = () => {
                const storedCategories = localStorage.getItem('categories');
                if (storedCategories) {
                    return JSON.parse(storedCategories);
                } else {
                    // Return the default list if nothing is in storage
                    return [
                        "バイト給料", "給料", "日本学生支援機構給付金", "川崎市大学等進学奨学金",
                        "千文基金", "篠原財団", "家賃", "初期費用", "住宅設定費", "家電・家具",
                        "光熱費", "電気ガス水道費", "インターネット", "NHK通信", "NHK通信料",
                        "学費", "学費減免", "教材費", "大学進学等自立生活支度費",
                        "専門学校入学支度金", "社会的自立支援費（私立）", "定期", "バス 定期",
                        "電車 定期", "交通費定期", "食費", "生活消耗品", "国民健康保険",
                        "娯楽費", "その他"
                    ].sort();
                }
            };

            const calculateAllBalances = () => {
                let previousMonthFinalBalance = 0;
                let grandTotalIncome = 0;
                let grandTotalExpenditure = 0;
                let firstDataRow = null;
                let lastDataRow = null;
                let isFirstMonthEver = true;
                let initialBalance = 0;

                state.financialData.forEach(year => {
                    year.months.sort((a, b) => {
                        const monthA = parseInt(a.month);
                        const monthB = parseInt(b.month);

                        if (year.year.endsWith('年度') || year.year === '1年次' || year.year === '2年次') {
                            let adjustedMonthA = monthA < 4 ? monthA + 12 : monthA;
                            let adjustedMonthB = monthB < 4 ? monthB + 12 : monthB;

                            if (year.year === '1年次') {
                                if (monthA === 3 && a.note.includes('貯金')) adjustedMonthA = 3;
                                if (monthB === 3 && b.note.includes('貯金')) adjustedMonthB = 3;
                            }

                            return adjustedMonthA - adjustedMonthB;
                        }

                        return monthA - monthB;
                    });

                    let yearTotalIncome = 0;
                    let yearTotalExpenditure = 0;

                    year.months.forEach(month => {
                        month.totalIncome = month.income.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);
                        month.totalExpenditure = month.expenditure.reduce((sum, item) => sum + (Number(item.amount) || 0), 0);

                        if (isFirstMonthEver) {
                            initialBalance = month.startingBalance;
                            previousMonthFinalBalance = month.startingBalance; // 最初の月のstartingBalanceで初期化
                            isFirstMonthEver = false;
                        } else {
                            month.startingBalance = previousMonthFinalBalance;
                        }

                        month.finalBalance = month.startingBalance + month.totalIncome - month.totalExpenditure;
                        previousMonthFinalBalance = month.finalBalance;

                        yearTotalIncome += month.totalIncome;
                        yearTotalExpenditure += month.totalExpenditure;
                    });

                    year.totalIncome = yearTotalIncome;
                    year.totalExpenditure = yearTotalExpenditure;
                    if (year.months.length > 0) {
                        year.finalBalance = year.months[year.months.length - 1].finalBalance;
                    }

                    grandTotalIncome += yearTotalIncome;
                    grandTotalExpenditure += yearTotalExpenditure;
                });

                state.totals = {
                    income: grandTotalIncome,
                    expenditure: grandTotalExpenditure,
                    balance: previousMonthFinalBalance,
                    initial: initialBalance
                };
            };

            // --- Modals ---
            const showModal = (modalElement) => {
                document.body.classList.add('modal-open');
                modalElement.style.display = 'flex';
            };
            const hideModal = (modalElement) => {
                document.body.classList.remove('modal-open');
                modalElement.style.display = 'none';
            };

            const showEditModal = (title, name = '', amount = '') => {
                modalTitle.textContent = title;
                itemNameInput.value = name;
                itemAmountInput.value = amount;

                // Populate datalist with existing item names
                const itemNamesDatalist = document.getElementById('item-names');
                itemNamesDatalist.innerHTML = '';
                state.categories.forEach(name => {
                    const option = document.createElement('option');
                    option.value = name;
                    itemNamesDatalist.appendChild(option);
                });

                showModal(modal);
            };

            // --- CRUD & Data Functions ---
            const handleAddItem = (yearId, monthId, type) => {
                state.editingItem = { yearId, monthId, type };
                showEditModal(`新しい${type === 'income' ? '収入' : '支出'}を追加`, '', 0);
            };

            const handleAddYear = (options) => {
                const { name, type, selectedMonths, afterYearId } = options;

                const lastMonthOfPreviousYear = findYear(afterYearId)?.months.slice(-1)[0];
                const startingBalance = lastMonthOfPreviousYear ? lastMonthOfPreviousYear.finalBalance : 0;

                const newYear = {
                    id: `year_${Date.now()}`,
                    year: name,
                    months: selectedMonths.map((monthName, index) => ({
                        id: `month_${Date.now()}_${index}`,
                        month: monthName,
                        note: '',
                        startingBalance: index === 0 ? startingBalance : 0,
                        income: [],
                        expenditure: []
                    }))
                };

                const afterIndex = state.financialData.findIndex(y => y.id === afterYearId);
                if (afterIndex > -1) {
                    state.financialData.splice(afterIndex + 1, 0, newYear);
                } else {
                    state.financialData.push(newYear);
                }

                rerender();
            };

            const handleEditItem = (yearId, monthId, type, itemId) => {
                const item = findItem(yearId, monthId, type, itemId);
                state.editingItem = { yearId, monthId, type, itemId };
                showEditModal(`${type === 'income' ? '収入' : '支出'}を編集`, item.name, item.amount || 0);
            };

            const handleDeleteItem = (yearId, monthId, type, itemId) => {
                if (!confirm('この項目を削除しますか？')) return;
                const month = findMonth(yearId, monthId);
                month[type] = month[type].filter(i => i.id !== itemId);
                rerender();
            };

            const handleYearNameUpdate = (yearId, newName) => {
                const year = findYear(yearId);
                if (year) {
                    year.year = newName;
                    saveState();
                }
            };

            const handleNoteUpdate = (yearId, monthId, newNote) => {
                const month = findMonth(yearId, monthId);
                month.note = newNote;
                saveState(); // No need to rerender, just save
            };

            const findYear = (yearId) => state.financialData.find(y => y.id === yearId);
            const findMonth = (yearId, monthId) => findYear(yearId)?.months.find(m => m.id === monthId);
            const findItem = (yearId, monthId, type, itemId) => findMonth(yearId, monthId)?.[type].find(i => i.id === itemId);

            const extractCategoriesFromFinancialData = () => {
                const allItems = new Set();
                state.financialData.forEach(year => {
                    year.months.forEach(month => {
                        month.income.forEach(item => allItems.add(item.name));
                        month.expenditure.forEach(item => allItems.add(item.name));
                    });
                });
                return Array.from(allItems).sort();
            };

            // --- New Data Creation ---
            const createNewFinancialData = (options) => {
                const { type, yearName, initialBalance } = options;
                const newYear = {
                    id: `year_${Date.now()}`,
                    year: yearName,
                    months: []
                };

                let months = [];
                if (type === 'fiscal') {
                    months = ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月'];
                } else { // calendar
                    months = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
                }

                newYear.months = months.map((monthName, index) => {
                    return {
                        id: `month_${Date.now()}_${index}`,
                        month: monthName,
                        note: '',
                        startingBalance: index === 0 ? Number(initialBalance) : 0,
                        income: [],
                        expenditure: []
                    };
                });

                return [newYear];
            };

            // --- Summary & Export/Import ---
            const generateSummaryData = () => {
                const summary = {};
                state.financialData.forEach(year => {
                    const yearSummary = { income: {}, expenditure: {} };
                    year.months.forEach(month => {
                        month.income.forEach(item => {
                            if (item.amount > 0) yearSummary.income[item.name] = (yearSummary.income[item.name] || 0) + item.amount;
                        });
                        month.expenditure.forEach(item => {
                            if (item.amount > 0) yearSummary.expenditure[item.name] = (yearSummary.expenditure[item.name] || 0) + item.amount;
                        });
                    });
                    summary[year.year] = yearSummary;
                });
                return summary;
            };

            const renderSummaryModal = () => {
                const summaryData = generateSummaryData();
                const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
                let content = '';

                for (const yearName in summaryData) {
                    const yearSummary = summaryData[yearName];
                    const incomeEntries = Object.entries(yearSummary.income).sort((a, b) => b[1] - a[1]);
                    const expenditureEntries = Object.entries(yearSummary.expenditure).sort((a, b) => b[1] - a[1]);
                    const totalIncome = incomeEntries.reduce((sum, [, amount]) => sum + amount, 0);
                    const totalExpenditure = expenditureEntries.reduce((sum, [, amount]) => sum + amount, 0);

                    content += `
                <div class="summary-year-section">
                    <h3 class="summary-year-title">${yearName}</h3>
                    <div class="summary-tables">
                        <div class="income">
                            <h4>収入</h4>
                            <table class="summary-table">
                                <thead><tr><th>項目</th><th class="amount">合計</th></tr></thead>
                                <tbody>${incomeEntries.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                                <tfoot><tr><td>合計</td><td class="amount">${formatter.format(totalIncome)}</td></tr></tfoot>
                            </table>
                        </div>
                        <div class="expenditure">
                            <h4>支出</h4>
                            <table class="summary-table">
                                <thead><tr><th>項目</th><th class="amount">合計</th></tr></thead>
                                <tbody>${expenditureEntries.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                                <tfoot><tr><td>合計</td><td class="amount">${formatter.format(totalExpenditure)}</td></tr></tfoot>
                            </table>
                        </div>
                    </div>
                </div>`;
                }
                summaryContent.innerHTML = content;
                showModal(summaryModal);
            };

            const renderCategorySummaryModal = () => {
                const incomeTotals = {};
                const expenditureTotals = {};

                state.financialData.forEach(year => {
                    year.months.forEach(month => {
                        month.income.forEach(item => {
                            if (item.amount > 0) {
                                incomeTotals[item.name] = (incomeTotals[item.name] || 0) + item.amount;
                            }
                        });
                        month.expenditure.forEach(item => {
                            if (item.amount > 0) {
                                expenditureTotals[item.name] = (expenditureTotals[item.name] || 0) + item.amount;
                            }
                        });
                    });
                });

                const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
                const sortedIncome = Object.entries(incomeTotals).sort((a, b) => b[1] - a[1]);
                const sortedExpenditure = Object.entries(expenditureTotals).sort((a, b) => b[1] - a[1]);

                const totalIncome = sortedIncome.reduce((sum, [, amount]) => sum + amount, 0);
                const totalExpenditure = sortedExpenditure.reduce((sum, [, amount]) => sum + amount, 0);

                const incomeTable = `
            <div class="income">
                <h4>収入合計</h4>
                <table class="summary-table">
                    <thead><tr><th>項目</th><th class="amount">合計金額</th></tr></thead>
                    <tbody>${sortedIncome.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                    <tfoot><tr><td>総合計</td><td class="amount">${formatter.format(totalIncome)}</td></tr></tfoot>
                </table>
            </div>
        `;
                const expenditureTable = `
            <div class="expenditure">
                <h4>支出合計</h4>
                <table class="summary-table">
                    <thead><tr><th>項目</th><th class="amount">合計金額</th></tr></thead>
                    <tbody>${sortedExpenditure.map(([name, amount]) => `<tr><td>${name}</td><td class="amount">${formatter.format(amount)}</td></tr>`).join('')}</tbody>
                    <tfoot><tr><td>総合計</td><td class="amount">${formatter.format(totalExpenditure)}</td></tr></tfoot>
                </table>
            </div>
        `;

                categorySummaryContent.innerHTML = `<div class="summary-tables">${incomeTable}${expenditureTable}</div>`;
                showModal(categorySummaryModal);
            };

            const EXCELJS_CDN_URL = "https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.3.0/exceljs.min.js";
            let excelJsLoadPromise = null;

            const ensureExcelJsLoaded = () => {
                if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
                if (excelJsLoadPromise) return excelJsLoadPromise;

                excelJsLoadPromise = new Promise((resolve, reject) => {
                    const script = document.createElement('script');
                    script.src = EXCELJS_CDN_URL;
                    script.async = true;
                    script.onload = () => resolve(window.ExcelJS);
                    script.onerror = () => {
                        excelJsLoadPromise = null;
                        reject(new Error('failed to load ExcelJS'));
                    };
                    document.head.appendChild(script);
                });

                return excelJsLoadPromise;
            };
            const exportToJson = () => {
                const dataToExport = {
                    userName: state.userName,
                    financialData: state.financialData
                };

                let exportObject;
                if (state.password) {
                    const encryptedData = CryptoJS.AES.encrypt(JSON.stringify(dataToExport), state.password).toString();
                    exportObject = {
                        isPasswordProtected: true,
                        data: encryptedData
                    };
                    showPopupMessage('データはパスワードで暗号化してエクスポートしました。');
                } else {
                    exportObject = {
                        isPasswordProtected: false,
                        data: dataToExport
                    };
                    showPopupMessage('データは暗号化せずにエクスポートしました。');
                }

                const outputData = JSON.stringify(exportObject, null, 2);
                const dataBlob = new Blob([outputData], { type: "application/json" });
                const url = URL.createObjectURL(dataBlob);
                const link = document.createElement("a");
                link.setAttribute("href", url);
                link.setAttribute("download", "financial_data.json");
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);
            };

            const exportToExcel = async () => {
                if (!state.financialData.length) {
                    showPopupMessage('エクスポートできるデータがありません。');
                    return;
                }

                let ExcelJS;
                try {
                    ExcelJS = await ensureExcelJsLoaded();
                } catch (err) {
                    console.error(err);
                    showPopupMessage('Excelライブラリの読み込みに失敗しました。インターネット接続を確認してもう一度お試しください。');
                    return;
                }

                if (!ExcelJS) {
                    showPopupMessage('Excelライブラリを初期化できなかったためエクスポートできません。');
                    return;
                }

                calculateAllBalances();

                const workbook = new ExcelJS.Workbook();
                const worksheet = workbook.addWorksheet('収支リスト', {
                    views: [{ state: 'frozen', xSplit: 0, ySplit: 4, showGridLines: false }]
                });

                worksheet.columns = [
                    { width: 12 }, { width: 20 }, { width: 20 },
                    { width: 20 }, { width: 20 },
                    { width: 18 }, { width: 18 }, { width: 24 },
                    { width: 20 }, { width: 20 }, { width: 24 }
                ];

                const applyBorder = (cell, variant = 'thin') => {
                    cell.border = ['top', 'left', 'bottom', 'right'].reduce((border, side) => {
                        border[side] = { style: variant, color: { argb: 'FFB0B0B0' } };
                        return border;
                    }, {});
                };

                const formatMoney = (row, indices = [3, 5, 6, 7]) => {
                    indices.forEach(idx => {
                        const cell = row.getCell(idx);
                        cell.numFmt = '"¥"#,##0';
                        cell.alignment = { horizontal: 'right' };
                    });
                };

                const exportOwnerName = (state.userName && state.userName.trim()) ? state.userName.trim() : '名前未設定';
                const workbookTitle = `${exportOwnerName} さんの収支ブック`;
                const subtitle = `作成: ${new Date().toLocaleString('ja-JP')}`;

                const titleRow = worksheet.addRow([workbookTitle]);
                worksheet.mergeCells(titleRow.number, 1, titleRow.number, 8);
                titleRow.font = { bold: true, size: 18, color: { argb: 'FF0F3A78' } };
                titleRow.alignment = { horizontal: 'center' };

                const subtitleRow = worksheet.addRow([subtitle]);
                worksheet.mergeCells(subtitleRow.number, 1, subtitleRow.number, 8);
                subtitleRow.font = { size: 11, color: { argb: 'FF6B7280' } };
                subtitleRow.alignment = { horizontal: 'center' };
                worksheet.addRow([]);

                const toNumber = (value) => {
                    const num = Number(value);
                    return Number.isFinite(num) ? num : 0;
                };

                let grandTotalIncome = 0;
                let grandTotalExpenditure = 0;
                let tableIndex = 1;
                const detailHeaders = ['月', '収入項目', '収入金額', '支出項目', '支出金額', '月次収支', '当月残額', 'メモ'];
                const summaryHeaders = ['月', '前月残額', '収入合計', '支出合計', '月次収支', '当月残額'];

                state.financialData.forEach(year => {
                    worksheet.addRow([]);
                    const yearRow = worksheet.addRow([`${year.year}`]);
                    worksheet.mergeCells(yearRow.number, 1, yearRow.number, 8);
                    yearRow.font = { bold: true, size: 13, color: { argb: 'FF0B5394' } };
                    yearRow.alignment = { horizontal: 'left' };
                    yearRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F0FE' } };
                    yearRow.eachCell(cell => applyBorder(cell, 'thin'));
                    worksheet.addRow([]);

                    const safeYearName = `${year.year}`.replace(/[^A-Za-z0-9]/g, '_') || 'Year';
                    const summaryStartRow = worksheet.lastRow.number + 1;
                    const summaryRows = [];
                    year.months.forEach(month => {
                        const monthIncomeTotal = toNumber(month.totalIncome || 0);
                        const monthExpenditureTotal = toNumber(month.totalExpenditure || 0);
                        const monthlyNet = toNumber(monthIncomeTotal - monthExpenditureTotal);
                        const monthFinalBalance = toNumber(month.finalBalance || 0);
                        summaryRows.push([
                            month.month || '',
                            toNumber(month.startingBalance || 0),
                            monthIncomeTotal,
                            monthExpenditureTotal,
                            monthlyNet,
                            monthFinalBalance
                        ]);
                    });

                    if (summaryRows.length === 0) {
                        summaryRows.push(['', 0, 0, 0, 0, 0]);
                    }

                    const summaryTableName = `Summary_${safeYearName}_${tableIndex++}`;
                    worksheet.addTable({
                        name: summaryTableName,
                        ref: `A${summaryStartRow}`,
                        headerRow: true,
                        style: { theme: 'TableStyleLight16', showRowStripes: true },
                        columns: summaryHeaders.map(name => ({ name })),
                        rows: summaryRows
                    });

                    const summaryFirstDataRow = summaryStartRow + 1;
                    summaryRows.forEach((_, idx) => {
                        const row = worksheet.getRow(summaryFirstDataRow + idx);
                        formatMoney(row, [2, 3, 4, 5, 6]);
                        row.getCell(1).alignment = { horizontal: 'center' };
                    });
                    worksheet.addRow([]);

                    const tableStartRow = worksheet.lastRow.number + 1;
                    const tableRows = [];
                    let yearIncomeTotal = 0;
                    let yearExpenditureTotal = 0;

                    year.months.forEach(month => {
                        const incomeItems = month.income.length ? month.income : [{ name: 'ー', amount: 0 }];
                        const expenditureItems = month.expenditure.length ? month.expenditure : [{ name: 'ー', amount: 0 }];
                        const rowCount = Math.max(incomeItems.length, expenditureItems.length);
                        const monthIncomeTotal = toNumber(month.totalIncome || 0);
                        const monthExpenditureTotal = toNumber(month.totalExpenditure || 0);
                        const monthlyNet = toNumber(monthIncomeTotal - monthExpenditureTotal);
                        const monthFinalBalance = toNumber(month.finalBalance || 0);
                        yearIncomeTotal += monthIncomeTotal;
                        yearExpenditureTotal += monthExpenditureTotal;

                        for (let i = 0; i < rowCount; i++) {
                            const incomeItem = incomeItems[i] || {};
                            const expenditureItem = expenditureItems[i] || {};
                            tableRows.push([
                                i === 0 ? (month.month || '') : '',
                                incomeItem.name || '',
                                toNumber(incomeItem.amount),
                                expenditureItem.name || '',
                                toNumber(expenditureItem.amount),
                                i === 0 ? monthlyNet : '',
                                i === 0 ? monthFinalBalance : '',
                                i === 0 ? (month.note || '') : ''
                            ]);
                        }
                    });

                    if (tableRows.length === 0) {
                        tableRows.push(['', '', 0, '', 0, 0, 0, '']);
                    }

                    const tableName = `Detail_${safeYearName}_${tableIndex++}`;
                    worksheet.addTable({
                        name: tableName,
                        ref: `A${tableStartRow}`,
                        headerRow: true,
                        style: { theme: 'TableStyleMedium6', showRowStripes: true, showFirstColumn: false, showLastColumn: false },
                        columns: detailHeaders.map(name => ({ name })),
                        rows: tableRows
                    });

                    const firstDataRow = tableStartRow + 1;
                    tableRows.forEach((_, idx) => {
                        const row = worksheet.getRow(firstDataRow + idx);
                        formatMoney(row);
                        row.getCell(1).alignment = { vertical: 'middle', horizontal: 'center' };
                        row.getCell(2).alignment = { horizontal: 'left' };
                        row.getCell(4).alignment = { horizontal: 'left' };
                        row.getCell(8).alignment = { horizontal: 'left', wrapText: true };
                    });

                    const yearTotalRow = worksheet.addRow([
                        '',
                        '年合計',
                        yearIncomeTotal,
                        '',
                        yearExpenditureTotal,
                        toNumber(yearIncomeTotal - yearExpenditureTotal),
                        '',
                        ''
                    ]);
                    yearTotalRow.font = { bold: true };
                    formatMoney(yearTotalRow);
                    yearTotalRow.eachCell(cell => applyBorder(cell, 'medium'));
                    yearTotalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
                    worksheet.addRow([]);

                    grandTotalIncome += yearIncomeTotal;
                    grandTotalExpenditure += yearExpenditureTotal;
                });

                const summaryRow = worksheet.addRow([
                    '',
                    '全期間合計',
                    grandTotalIncome,
                    '',
                    grandTotalExpenditure,
                    toNumber(grandTotalIncome - grandTotalExpenditure),
                    '',
                    ''
                ]);
                summaryRow.font = { bold: true, size: 12 };
                formatMoney(summaryRow);
                summaryRow.eachCell(cell => applyBorder(cell, 'medium'));
                summaryRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF5FF' } };

                const buffer = await workbook.xlsx.writeBuffer();
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `financial_${new Date().toISOString().slice(0, 10)}.xlsx`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);

                showPopupMessage('Excelファイルをエクスポートしました。');
            };

            // --- Word Export (docx) ---
            // Use browser-ready UMD bundles and fallbacks
            const DOCX_CDN_URLS = [
                "https://cdnjs.cloudflare.com/ajax/libs/docx/8.3.2/docx.umd.min.js",
                "https://cdn.jsdelivr.net/npm/docx@8.3.2/build/index.js",
                "https://unpkg.com/docx@8.3.2/build/index.js"
            ];
            let docxLoadPromise = null;

            const loadDocxFrom = (url) => new Promise((resolve, reject) => {
                const script = document.createElement('script');
                script.src = url;
                script.async = true;
                script.crossOrigin = 'anonymous';
                script.onload = () => resolve(window.docx);
                script.onerror = () => reject(new Error(`failed to load docx from ${url}`));
                document.head.appendChild(script);
            });

            const ensureDocxLoaded = () => {
                if (window.docx) return Promise.resolve(window.docx);
                if (docxLoadPromise) return docxLoadPromise;

                docxLoadPromise = DOCX_CDN_URLS.reduce((chain, url) => {
                    return chain.catch(() => loadDocxFrom(url));
                }, Promise.reject());

                // If all failed, reset so we can retry later
                docxLoadPromise = docxLoadPromise.catch(err => {
                    docxLoadPromise = null;
                    throw err;
                });

                return docxLoadPromise;
            };

            const exportToWord = async () => {
                if (!state.financialData.length) {
                    showPopupMessage('エクスポートできるデータがありません。');
                    return;
                }

                let docx;
                try {
                    docx = await ensureDocxLoaded();
                } catch (err) {
                    console.error(err);
                    showPopupMessage('Wordライブラリの読み込みに失敗しました。インターネット接続を確認してもう一度お試しください。');
                    return;
                }

                if (!docx || !docx.Document || !docx.Packer) {
                    showPopupMessage('Wordライブラリを初期化できなかったためエクスポートできません。');
                    return;
                }

                calculateAllBalances();

                const { Document, Paragraph, TextRun, Table, TableRow, TableCell, HeadingLevel, WidthType, AlignmentType, Packer, BorderStyle } = docx;
                const owner = (state.userName && state.userName.trim()) ? state.userName.trim() : '名前未設定';
                const docTitle = `${owner} さんの収支レポート`;
                const docSubtitle = `作成: ${new Date().toLocaleString('ja-JP')}`;
                const formatCurrency = (n) => new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' }).format(Number.isFinite(n) ? n : 0);

                const makeHeading = (text, level = HeadingLevel.HEADING_2, after = 200) =>
                    new Paragraph({ text, heading: level, spacing: { after } });

                const makeTable = (headers, rows, widthPercent = 100) => {
                    const headerCells = headers.map(h => new TableCell({
                        children: [new Paragraph({ text: h, bold: true })],
                        shading: { fill: 'E8F0FE' }
                    }));

                    const bodyRows = rows.map(r => new TableRow({
                        children: r.map(cell => new TableCell({
                            children: [new Paragraph({ text: cell })]
                        }))
                    }));

                    return new Table({
                        width: { size: widthPercent * 50, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({ children: headerCells }),
                            ...bodyRows
                        ],
                        margins: { top: 80, bottom: 80, left: 80, right: 80 },
                        borders: {
                            top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                            bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                            left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                            right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
                            insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' },
                            insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'DDDDDD' }
                        }
                    });
                };

                const doc = new Document({
                    sections: [{
                        children: [
                            new Paragraph({
                                children: [new TextRun({ text: docTitle, size: 32, bold: true, color: '0F3A78' })],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 160 }
                            }),
                            new Paragraph({
                                children: [new TextRun({ text: docSubtitle, size: 22, color: '666666' })],
                                alignment: AlignmentType.CENTER,
                                spacing: { after: 260 }
                            })
                        ]
                    }]
                });

                let totalIncomeAll = 0;
                let totalExpenditureAll = 0;

                state.financialData.forEach(year => {
                    const children = [];
                    children.push(makeHeading(`${year.year}`, HeadingLevel.HEADING_2, 150));
                    children.push(makeHeading('月次サマリー', HeadingLevel.HEADING_3, 120));

                    const summaryRows = year.months.map(m => {
                        const income = Number(m.totalIncome) || 0;
                        const exp = Number(m.totalExpenditure) || 0;
                        const net = income - exp;
                        totalIncomeAll += income;
                        totalExpenditureAll += exp;
                        return [
                            m.month || '',
                            formatCurrency(m.startingBalance || 0),
                            formatCurrency(income),
                            formatCurrency(exp),
                            formatCurrency(net),
                            formatCurrency(m.finalBalance || 0)
                        ];
                    });

                    if (summaryRows.length === 0) summaryRows.push(['', formatCurrency(0), formatCurrency(0), formatCurrency(0), formatCurrency(0), formatCurrency(0)]);
                    children.push(makeTable(['月', '前月残額', '収入合計', '支出合計', '月次収支', '当月残額'], summaryRows, 100));
                    children.push(new Paragraph({ spacing: { after: 200 } }));

                    children.push(makeHeading('詳細リスト', HeadingLevel.HEADING_3, 120));
                    year.months.forEach(m => {
                        const monthTitle = `${m.month || ''} の内訳`;
                        children.push(makeHeading(monthTitle, HeadingLevel.HEADING_4, 80));

                        const detailRows = [];
                        const maxRows = Math.max(m.income.length || 0, m.expenditure.length || 0, 1);
                        for (let i = 0; i < maxRows; i++) {
                            const inc = m.income[i] || {};
                            const exp = m.expenditure[i] || {};
                            detailRows.push([
                                i === 0 ? (m.month || '') : '',
                                inc.name || '-',
                                formatCurrency(inc.amount || 0),
                                exp.name || '-',
                                formatCurrency(exp.amount || 0),
                                i === 0 ? formatCurrency((Number(m.totalIncome) || 0) - (Number(m.totalExpenditure) || 0)) : '',
                                i === 0 ? formatCurrency(m.finalBalance || 0) : ''
                            ]);
                        }

                        children.push(makeTable(
                            ['月', '収入項目', '収入金額', '支出項目', '支出金額', '月次収支', '当月残額'],
                            detailRows,
                            100
                        ));

                        if (m.note) {
                            children.push(new Paragraph({
                                children: [
                                    new TextRun({ text: 'メモ: ', bold: true }),
                                    new TextRun({ text: m.note })
                                ],
                                spacing: { after: 180 }
                            }));
                        } else {
                            children.push(new Paragraph({ spacing: { after: 120 } }));
                        }
                    });

                    doc.addSection({ children });
                });

                const netAll = totalIncomeAll - totalExpenditureAll;
                doc.addSection({
                    children: [
                        makeHeading('全期間サマリー', HeadingLevel.HEADING_2, 120),
                        makeTable(
                            ['総収入', '総支出', '純収支'],
                            [[formatCurrency(totalIncomeAll), formatCurrency(totalExpenditureAll), formatCurrency(netAll)]],
                            70
                        )
                    ]
                });

                const blob = await Packer.toBlob(doc);
                const url = URL.createObjectURL(blob);
                const link = document.createElement('a');
                link.href = url;
                link.download = `financial_${new Date().toISOString().slice(0, 10)}.docx`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(url);

                showPopupMessage('Wordファイルをエクスポートしました。');
            };

            const sanitizeText = (input, maxLen = 5000) => {
                if (typeof input !== 'string') return '';
                return input.slice(0, maxLen);
            };

            const sanitizeNumber = (input) => {
                const num = Number(input);
                return Number.isFinite(num) ? num : 0;
            };

            const sanitizeFinancialData = (raw) => {
                if (!Array.isArray(raw)) return [];
                return raw.map(year => {
                    const yearId = year?.id ? String(year.id) : `year_${Date.now()}_${Math.random()}`;
                    const months = Array.isArray(year?.months) ? year.months : [];
                    return {
                        id: yearId,
                        year: sanitizeText(year?.year || ''),
                        months: months.map(month => {
                            const monthId = month?.id ? String(month.id) : `month_${Date.now()}_${Math.random()}`;
                            const income = Array.isArray(month?.income) ? month.income : [];
                            const expenditure = Array.isArray(month?.expenditure) ? month.expenditure : [];
                            return {
                                id: monthId,
                                month: sanitizeText(month?.month || ''),
                                note: sanitizeText(month?.note || ''),
                                startingBalance: sanitizeNumber(month?.startingBalance || 0),
                                income: income.map(item => ({
                                    id: item?.id ? String(item.id) : `inc_${Date.now()}_${Math.random()}`,
                                    name: sanitizeText(item?.name || ''),
                                    amount: sanitizeNumber(item?.amount || 0)
                                })),
                                expenditure: expenditure.map(item => ({
                                    id: item?.id ? String(item.id) : `exp_${Date.now()}_${Math.random()}`,
                                    name: sanitizeText(item?.name || ''),
                                    amount: sanitizeNumber(item?.amount || 0)
                                }))
                            };
                        })
                    };
                });
            };

            const importData = (data) => {
                if (!confirm('現在のデータをインポートしたデータで上書きしますか？この操作は元に戻せません。')) return;

                const sanitized = data?.financialData ? sanitizeFinancialData(data.financialData) : sanitizeFinancialData(data);
                if (!Array.isArray(sanitized) || sanitized.length === 0) {
                    showPopupMessage('インポートデータの形式が不正です。');
                    return;
                }

                state.financialData = processData(sanitized);
                state.userName = data?.userName ? sanitizeText(data.userName, 200) : '';
                userNameDisplay.textContent = state.userName ? `${state.userName}の収支` : '';
                rerender();
                showPopupMessage('データを正常にインポートしました。');
            };

            const handleImportedFileContent = (content) => {
                try {
                    let importObject;
                    try {
                        importObject = JSON.parse(content);
                    } catch (e) {
                        showPopupMessage('Invalid or encrypted file. Legacy fixed-key imports are disabled.');
                        console.error('JSON import failed:', e);
                        return;
                    }


                    if (importObject.isPasswordProtected) {
                        const password = prompt('このファイルはパスワードで保護されています。パスワードを入力してください:');
                        if (password === null) return; // キャンセルされた

                        try {
                            const bytes = CryptoJS.AES.decrypt(importObject.data, password);
                            const decryptedString = bytes.toString(CryptoJS.enc.Utf8);
                            if (!decryptedString) throw new Error("Decryption failed");
                            const decryptedData = JSON.parse(decryptedString);
                            importData(decryptedData);
                        } catch (err) {
                            showPopupMessage('パスワードが違うか、ファイルが破損しています。');
                            console.error('復号化に失敗:', err);
                        }
                    } else {
                        // isPasswordProtected: false またはフラグなしの新しい形式
                        importData(importObject.data || importObject); // 古い形式も考慮
                    }
                } catch (err) {
                    showPopupMessage('無効なファイル形式です。');
                    console.error('JSONのインポートに失敗しました:', err);
                }
            };

            const importFromJson = () => {
                const input = document.createElement('input');
                input.type = 'file';
                input.accept = 'application/json';
                input.onchange = e => {
                    const file = e.target.files[0];
                    if (!file) return;
                    const reader = new FileReader();
                    reader.onload = readerEvent => {
                        handleImportedFileContent(readerEvent.target.result);
                    };
                    reader.readAsText(file);
                };
                input.click();
            };

            // --- Drag & Drop for Import ---
            appContainer.addEventListener('dragover', (e) => {
                e.preventDefault();
                appContainer.classList.add('drag-over');
            });

            appContainer.addEventListener('dragleave', (e) => {
                e.preventDefault();
                appContainer.classList.remove('drag-over');
            });

            appContainer.addEventListener('drop', (e) => {
                e.preventDefault();
                appContainer.classList.remove('drag-over');

                const files = e.dataTransfer.files;
                if (files.length > 0) {
                    const file = files[0];
                    if (file.type === 'application/json') {
                        const reader = new FileReader();
                        reader.onload = readerEvent => {
                            handleImportedFileContent(readerEvent.target.result);
                        };
                        reader.readAsText(file);
                    } else {
                        showPopupMessage('JSONファイルのみをドロップしてください。');
                    }
                }
            });

            // --- Rendering ---
            const escapeHTML = (value) => {
                const text = value === undefined || value === null ? '' : String(value);
                return text
                    .replace(/&/g, '&amp;')
                    .replace(/</g, '&lt;')
                    .replace(/>/g, '&gt;')
                    .replace(/"/g, '&quot;')
                    .replace(/'/g, '&#39;');
            };

            const highlightSearchTerm = (text, searchTerm) => {
                const safeText = escapeHTML(text ?? '');
                if (!searchTerm) return safeText;
                const escapedSearchTerm = searchTerm.replace(/[.*+?^${}()|[\\]/g, '\\$&');
                const regex = new RegExp(`(${escapedSearchTerm})`, 'gi');
                return safeText.replace(regex, '<span class="search-highlight">$1</span>');
            };

            const render = () => {
                if (!appContainer) return;

                calculateAllBalances();
                const formatter = new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' });
                appContainer.innerHTML = '';
                const searchTerm = (state.searchTerm || '').toLowerCase();

                const grandTotalIncome = state.totals.income;
                const grandTotalExpenditure = state.totals.expenditure;
                const netTotal = state.totals.balance;
                const initialBalance = state.totals.initial;

                appContainer.innerHTML = `
        <div class="total-summary-card">
            <h2>総合収支</h2>
            <div class="summary-grid">
                <div class="summary-grid-item">
                    <div class="label">初期残高</div>
                    ${state.isEditMode
                        ? `<input type="number" class="initial-balance-editor" value="${initialBalance}">`
                        : `<div class="value">${formatter.format(initialBalance)}</div>`}
                </div>
                <div class="summary-grid-item"><div class="label">総収入</div><div class="value income">${formatter.format(grandTotalIncome)}</div></div>
                <div class="summary-grid-item"><div class="label">総支出</div><div class="value expenditure">${formatter.format(grandTotalExpenditure)}</div></div>
                <div class="summary-grid-item"><div class="label">最終残高</div><div class="value net">${formatter.format(netTotal)}</div></div>
            </div>
        </div>`;

                const filteredData = state.financialData.map(year => {
                    const filteredMonths = year.months.filter(month => {
                        if (!searchTerm) return true;
                        const inMonth = month.month.toLowerCase().includes(searchTerm);
                        const inNote = month.note.toLowerCase().includes(searchTerm);
                        const inIncome = month.income.some(i => i.name.toLowerCase().includes(searchTerm));
                        const inExpenditure = month.expenditure.some(e => e.name.toLowerCase().includes(searchTerm));
                        return inMonth || inNote || inIncome || inExpenditure;
                    });
                    return { ...year, months: filteredMonths };
                }).filter(year => year.months.length > 0);

                if (filteredData.length === 0) {
                    appContainer.innerHTML += '<div class="empty-list-placeholder">データがありません。</div>';
                    return;
                }

                filteredData.forEach(year => {
                    const yearSection = document.createElement('div');
                    yearSection.className = 'year-section';

                    const yearHeaderContent = state.isEditMode
                        ? `<div class="year-header-editor">
                       <input type="text" class="year-title-editor" value="${year.year}" data-year-id="${year.id}">
                       <button class="control-btn edit-months-btn" data-year-id="${year.id}" title="月の編集">${icons.edit}</button>
                       <button class="control-btn add-year-btn" data-after-year-id="${year.id}" title="新しい年度/年を追加">${icons.add}</button>
                   </div>`
                        : `<h2 class="year-title">${highlightSearchTerm(year.year, searchTerm)}</h2>`;

                    yearSection.innerHTML = `<div class="year-header">${yearHeaderContent}</div><div class="months-grid"></div>`;
                    const monthsGrid = yearSection.querySelector('.months-grid');

                    year.months.forEach(month => {
                        const totalIncome = month.totalIncome;
                        const totalExpenditure = month.totalExpenditure;
                        const monthlyNet = totalIncome - totalExpenditure;
                        const incomePercent = (totalIncome + totalExpenditure) > 0 ? (totalIncome / (totalIncome + totalExpenditure)) * 100 : 0;

                        const incomeItems = month.income.length > 0 ? month.income.map(item => `
                    <li data-year-id="${year.id}" data-month-id="${month.id}" data-item-id="${item.id}" data-type="income">
                        <span>${highlightSearchTerm(item.name, searchTerm)}</span><span>${formatter.format(item.amount)}</span>
                        <div class="edit-controls"><button class="control-btn edit-btn">${icons.edit}</button><button class="control-btn delete-btn">${icons.delete}</button></div>
                    </li>`).join('') : '<div class="empty-list-placeholder">収入データがありません</div>';

                        const expenditureItems = month.expenditure.length > 0 ? month.expenditure.map(item => `
                     <li data-year-id="${year.id}" data-month-id="${month.id}" data-item-id="${item.id}" data-type="expenditure">
                        <span>${highlightSearchTerm(item.name, searchTerm)}</span><span>${formatter.format(item.amount)}</span>
                        <div class="edit-controls"><button class="control-btn edit-btn">${icons.edit}</button><button class="control-btn delete-btn">${icons.delete}</button></div>
                    </li>`).join('') : '<div class="empty-list-placeholder">支出データがありません</div>';

                        const card = document.createElement('div');
                        card.className = 'month-card';
                        if (monthlyNet > 0) card.classList.add('positive-flow');
                        if (monthlyNet < 0) card.classList.add('negative-flow');

                        card.innerHTML = `
                    <div class="month-header">
                        <h3 class="month-title">${highlightSearchTerm(month.month, searchTerm)}</h3>
                    </div>
                    <div class="summary">
                        <div class="summary-item"><span class="label">前月残額:</span> <span>${formatter.format(month.startingBalance)}</span></div>
                        <div class="summary-item"><span class="label">収入合計:</span> <span class="value income">${formatter.format(totalIncome)}</span></div>
                        <div class="summary-item"><span class="label">支出合計:</span> <span class="value expenditure">${formatter.format(totalExpenditure)}</span></div>
                    </div>
                    <div class="chart-container" data-year-id="${year.id}" data-month-id="${month.id}"><div class="chart"><div class="chart-income" style="width: ${incomePercent}%"></div><div class="chart-expenditure" style="width: ${100 - incomePercent}%"></div><div class="chart-label">収入${incomePercent.toFixed(0)}% / 支出${(100 - incomePercent).toFixed(0)}%</div></div></div>
                    <div class="details">
                        <div class="notes-section">
                            <h4>メモ</h4>
                            <div class="notes-content" style="display: ${state.isEditMode ? 'none' : 'block'};">${highlightSearchTerm(month.note || '-', searchTerm)}</div>
                            <textarea class="notes-editor" style="display: ${state.isEditMode ? 'block' : 'none'};" data-year-id="${year.id}" data-month-id="${month.id}">${month.note}</textarea>
                        </div>
                        <div class="income-section"><h4>収入 <div class="edit-controls"><button class="control-btn add-btn" data-year-id="${year.id}" data-month-id="${month.id}" data-type="income">${icons.add}</button></div></h4><ul>${incomeItems}</ul></div>
                        <div class="expenditure-section"><h4>支出 <div class="edit-controls"><button class="control-btn add-btn" data-year-id="${year.id}" data-month-id="${month.id}" data-type="expenditure">${icons.add}</button></div></h4><ul>${expenditureItems}</ul></div>
                    </div>
                    <div class="final-balance">最終残額: ${formatter.format(month.finalBalance)}</div>
                `;
                        monthsGrid.appendChild(card);
                    });
                    appContainer.appendChild(yearSection);
                });

                initTooltips();
            };

            const initTooltips = () => {
                tippy('.chart-container', {
                    content: (reference) => {
                        const { yearId, monthId } = reference.dataset;
                        const month = findMonth(yearId, monthId);
                        if (!month) return 'データが見つかりません';

                        const total = month.totalIncome + month.totalExpenditure;
                        if (total === 0) {
                            return '収入・支出がありません';
                        }

                        const incomePercent = (month.totalIncome / total) * 100;
                        const expenditurePercent = (month.totalExpenditure / total) * 100;

                        return `収入: ${incomePercent.toFixed(1)}%<br>支出: ${expenditurePercent.toFixed(1)}%`;
                    },
                    allowHTML: true,
                    theme: 'light',
                });
            };

            const rerender = () => {
                saveState();
                render();
            };

            // --- Event Listeners ---
            if (passwordForm) {
                passwordForm.addEventListener('submit', async (e) => {
                    e.preventDefault();
                    const enteredPassword = passwordInput.value;
                    passwordError.textContent = ''; // Clear previous errors
                    await loadAndRenderData(enteredPassword);
                });
            }

            if (newDataButton) {
                newDataButton.addEventListener('click', () => {
                    if (startYearInput) {
                        startYearInput.value = new Date().getFullYear();
                        if (document.querySelector('input[name="creation-type"]:checked').value === 'fiscal') {
                            startYearInput.value = '1年次';
                        }
                    }
                    if (initialBalanceInput) initialBalanceInput.value = 0;
                    if (newDataModal) showModal(newDataModal);
                });

                const creationTypeRadios = document.querySelectorAll('input[name="creation-type"]');
                creationTypeRadios.forEach(radio => {
                    radio.addEventListener('change', (e) => {
                        if (e.target.value === 'fiscal') {
                            startYearInput.value = '1年次';
                        } else {
                            startYearInput.value = new Date().getFullYear();
                        }
                    });
                });
            }

            if (newDataModal) {
                newDataCancelButton.addEventListener('click', () => hideModal(newDataModal));
                newDataCloseButton.addEventListener('click', () => hideModal(newDataModal));
                newDataModal.addEventListener('click', (e) => {
                    if (e.target === newDataModal) hideModal(newDataModal);
                });
            }

            if (newDataForm) {
                newDataForm.addEventListener('submit', (e) => {
                    e.preventDefault();
                    const creationType = document.querySelector('input[name="creation-type"]:checked').value;
                    const yearName = startYearInput.value.trim();
                    const initialBalance = parseInt(initialBalanceInput.value, 10);
                    const userName = userNameInput.value.trim();

                    if (!userName) {
                        showPopupMessage('名前を入力してください。');
                        return;
                    }
                    if (!yearName) {
                        showPopupMessage('有効な年/年度名を入力してください。');
                        return;
                    }
                    if (isNaN(initialBalance)) {
                        showPopupMessage('有効な初期残高を入力してください。');
                        return;
                    }

                    if (confirm('現在のデータをすべて削除し、新しいデータを作成します。よろしいですか？この操作は元に戻せません。')) {
                        const newData = createNewFinancialData({
                            type: creationType,
                            yearName: yearName,
                            initialBalance: initialBalance
                        });
                        state.financialData = processData(newData);
                        state.userName = userName;
                        userNameDisplay.textContent = `${state.userName}の収支`;
                        hideModal(newDataModal);
                        rerender();
                    }
                });
            }

            if (editModeToggle) {
                editModeToggle.addEventListener('click', () => {
                    state.isEditMode = !state.isEditMode;
                    document.body.classList.toggle('edit-mode', state.isEditMode);
                    editModeToggle.textContent = state.isEditMode ? '編集モード終了' : '編集モード';
                    editModeToggle.classList.toggle('active', state.isEditMode);

                    if (state.isEditMode) {
                        userNameDisplay.style.display = 'none';
                        userNameEdit.style.display = 'block';
                        userNameEdit.value = state.userName;
                        userNameEdit.focus();
                    } else {
                        userNameDisplay.style.display = 'block';
                        userNameEdit.style.display = 'none';
                        // 保存処理はchangeイベントで行う
                    }

                    rerender(); // Rerender to show/hide notes editor
                });
            }

            if (userNameEdit) {
                userNameEdit.addEventListener('change', (e) => {
                    const newName = e.target.value.trim();
                    if (newName) {
                        state.userName = newName;
                        userNameDisplay.textContent = `${state.userName}の収支`;
                        saveState();
                    } else {
                        e.target.value = state.userName; // 空の場合は元に戻す
                        showPopupMessage('名前は空にできません。');
                    }
                });
            }

            if (resetDataButton) {
                resetDataButton.addEventListener("click", async () => {
                    if (confirm("現在のすべてのデータをリセットしますか？この操作は元に戻せません。")) {
                        localStorage.removeItem("financialData");
                        localStorage.removeItem("userName");
                        localStorage.removeItem("password");
                        localStorage.removeItem("dataIsEncrypted");
                        state.userName = "";
                        state.password = "";
                        userNameDisplay.textContent = "";
                        state.financialData = await loadDefaultData();
                        showPopupMessage("データを初期状態にリセットしました。");
                        rerender();
                    }
                });
            }


            if (summaryButton) {
                summaryButton.addEventListener('click', renderSummaryModal);
            }

            if (summaryCloseButton) {
                summaryCloseButton.addEventListener('click', () => hideModal(summaryModal));
            }

            if (summaryModal) {
                summaryModal.addEventListener('click', (e) => {
                    if (e.target === summaryModal) hideModal(summaryModal);
                });
            }

            if (categorySummaryButton && categorySummaryModal && categorySummaryCloseButton) {
                categorySummaryButton.addEventListener('click', renderCategorySummaryModal);
                categorySummaryCloseButton.addEventListener('click', () => hideModal(categorySummaryModal));
                categorySummaryModal.addEventListener('click', (e) => {
                    if (e.target === categorySummaryModal) hideModal(categorySummaryModal);
                });
            }

            if (settingsButton) {
                settingsButton.addEventListener('click', () => {
                    // Populate categories
                    categoryInput.value = state.categories.join('\n');

                    // Populate password settings
                    passwordSettingInput.value = state.password; // Show current password
                    requirePasswordToggle.checked = state.requirePasswordOnLoad || !!state.password;
                    requirePasswordToggle.disabled = !!state.password;
                    requirePasswordToggleGroup.style.display = 'flex';

                    showModal(settingsModal);
                });
            }

            if (settingsCloseButton) {
                settingsCloseButton.addEventListener('click', () => hideModal(settingsModal));
            }

            if (settingsModal) {
                settingsModal.addEventListener('click', (e) => {
                    if (e.target === settingsModal) hideModal(settingsModal);
                });
            }

            if (cancelCategoriesButton) {
                cancelCategoriesButton.addEventListener('click', () => hideModal(settingsModal));
            }

            if (saveSettingsButton) {
                saveSettingsButton.addEventListener('click', () => {
                    const newCategories = categoryInput.value.split('\n').map(cat => cat.trim()).filter(cat => cat !== '');
                    state.categories = Array.from(new Set(newCategories)); // dedupe

                    const newPassword = passwordSettingInput.value.trim();
                    if (newPassword !== state.password) {
                        if (newPassword === '') {
                            if (state.password && !confirm('Removing the password will leave data unencrypted. Proceed?')) {
                                return;
                            }
                            state.password = '';
                            showPopupMessage('Password removed. Data will be stored unencrypted.');
                        } else {
                            state.password = newPassword;
                            showPopupMessage('Password set/updated.');
                        }
                    }

                    state.requirePasswordOnLoad = !!state.password || requirePasswordToggle.checked;
                    requirePasswordToggle.disabled = !!state.password;

                    hideModal(settingsModal);
                    rerender();
                    showPopupMessage('Settings saved.');
                });
            }

            if (autoGetCategoriesButton) {
                autoGetCategoriesButton.addEventListener('click', () => {
                    // Get existing categories from the textarea
                    const existingCategories = categoryInput.value.split('\n').map(cat => cat.trim()).filter(cat => cat);

                    // Get categories from financial data
                    const extractedCategories = extractCategoriesFromFinancialData();

                    // Combine them and get unique values
                    const combined = new Set([...existingCategories, ...extractedCategories]);

                    // Sort and update the textarea
                    categoryInput.value = Array.from(combined).sort().join('\n');
                });
            }

            if (exportJsonButton) exportJsonButton.addEventListener('click', exportToJson);
            if (exportExcelButton) exportExcelButton.addEventListener('click', exportToExcel);
            if (exportWordButton) exportWordButton.addEventListener('click', exportToWord);
            importJsonButton.addEventListener('click', importFromJson);

            searchBox.addEventListener('input', (e) => {
                state.searchTerm = e.target.value;
                render();
            });

            appContainer.addEventListener('click', e => {
                if (!state.isEditMode) return;
                const addBtn = e.target.closest('.add-btn');
                const editBtn = e.target.closest('.edit-btn');
                const deleteBtn = e.target.closest('.delete-btn');
                const addYearBtn = e.target.closest('.add-year-btn');

                if (addBtn) handleAddItem(addBtn.dataset.yearId, addBtn.dataset.monthId, addBtn.dataset.type);
                else if (editBtn) {
                    const li = e.target.closest('li');
                    handleEditItem(li.dataset.yearId, li.dataset.monthId, li.dataset.type, li.dataset.itemId);
                } else if (deleteBtn) {
                    const li = e.target.closest('li');
                    handleDeleteItem(li.dataset.yearId, li.dataset.monthId, li.dataset.type, li.dataset.itemId);
                } else if (addYearBtn) {
                    const afterYearId = addYearBtn.dataset.afterYearId;
                    openAddYearModal(afterYearId);
                } else if (e.target.closest('.edit-months-btn')) {
                    const editMonthsBtn = e.target.closest('.edit-months-btn');
                    const yearId = editMonthsBtn.dataset.yearId;
                    openEditMonthsModal(yearId);
                }
            });

            appContainer.addEventListener('change', e => {
                if (!state.isEditMode) return;

                if (e.target.classList.contains('notes-editor')) {
                    const { yearId, monthId } = e.target.dataset;
                    handleNoteUpdate(yearId, monthId, e.target.value);
                } else if (e.target.classList.contains('year-title-editor')) {
                    const { yearId } = e.target.dataset;
                    const newYearName = e.target.value.trim();
                    const year = findYear(yearId);

                    if (newYearName && year) {
                        handleYearNameUpdate(yearId, newYearName);
                    } else if (year) {
                        e.target.value = year.year; // Restore original value
                        showPopupMessage('年度名は空にできません。');
                    }
                } else if (e.target.classList.contains('initial-balance-editor')) {
                    const newInitialBalance = parseInt(e.target.value, 10);
                    if (!isNaN(newInitialBalance)) {
                        // 最初の月のstartingBalanceを更新
                        if (state.financialData.length > 0 && state.financialData[0].months.length > 0) {
                            state.financialData[0].months[0].startingBalance = newInitialBalance;
                            rerender();
                        }
                    } else {
                        showPopupMessage('有効な初期残高を入力してください。');
                        rerender(); // 無効な入力の場合は元の値を再表示
                    }
                }
            });

            modalForm.addEventListener('submit', e => {
                e.preventDefault();
                const { yearId, monthId, type, itemId } = state.editingItem;
                const name = itemNameInput.value.trim();
                const amount = parseInt(itemAmountInput.value, 10);
                if (!name || isNaN(amount)) return showPopupMessage('有効な項目名と金額を入力してください。');

                // Add new item name to categories if it doesn't exist
                if (!state.categories.includes(name)) {
                    state.categories.push(name);
                    state.categories.sort(); // ソートして表示順を整える
                }

                const month = findMonth(yearId, monthId);
                if (itemId) {
                    const item = month[type].find(i => i.id === itemId);
                    item.name = name; item.amount = amount;
                } else {
                    month[type].push({ id: `item_${Date.now()}`, name, amount });
                }
                hideModal(modal);
                rerender();
            });

            cancelButton.addEventListener('click', () => hideModal(modal));

            // --- Add Year Modal ---
            let currentAfterYearId = null;

            const populateMonthCheckboxes = (isFiscal) => {
                const months = isFiscal
                    ? ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
                    : ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];

                addYearMonthsContainer.innerHTML = months.map(month => `
            <div class="month-checkbox-item">
                <input type="checkbox" id="bulk-add-${month}" name="selectedMonths" value="${month}" checked>
                <label for="bulk-add-${month}">${month}月</label>
            </div>
        `).join('');
            };

            const openAddYearModal = (afterYearId) => {
                currentAfterYearId = afterYearId;
                const afterYear = findYear(afterYearId);
                const currentYearName = afterYear ? afterYear.year : '';

                let nextYearName = '';
                const nendoMatch = currentYearName.match(/(\d+)年次/);
                const yearMatch = currentYearName.match(/(\d{4})/);

                let isFiscalType = false;
                if (afterYear) {
                    isFiscalType = afterYear.year.includes('年度') || afterYear.year.includes('年次');
                }

                if (nendoMatch) {
                    const yearNumber = parseInt(nendoMatch[1], 10);
                    nextYearName = `${yearNumber + 1}年次`;
                } else if (yearMatch) {
                    const yearNumber = parseInt(yearMatch[1], 10);
                    nextYearName = currentYearName.replace(String(yearNumber), String(yearNumber + 1));
                } else {
                    nextYearName = isFiscalType ? `${new Date().getFullYear()}年度` : new Date().getFullYear().toString();
                }

                populateMonthCheckboxes(isFiscalType);
                newYearNameInput.value = nextYearName;

                // ラジオボタンのイベントリスナーと「新規作成」の表示/非表示制御
                const creationMethodRadios = document.querySelectorAll('input[name="creation-method"]');
                const copySourceContainer = document.getElementById('copy-source-container');
                const copySourceSelect = document.getElementById('copy-source-year');
                const newMonthsContainer = document.getElementById('new-months-container');
                const monthSelectionContainer = document.getElementById('month-selection-container');

                const handleCreationMethodChange = () => {
                    const selectedMethod = document.querySelector('input[name="creation-method"]:checked').value;
                    if (selectedMethod === 'copy') {
                        // コピーモードの場合
                        copySourceContainer.style.display = 'block';
                        newMonthsContainer.style.display = 'block';
                        monthSelectionContainer.style.display = 'none';
                        // コピー元選択肢を生成
                        populateCopySourceSelect();
                    } else {
                        // 新規作成モードの場合
                        copySourceContainer.style.display = 'none';
                        newMonthsContainer.style.display = 'block';
                        monthSelectionContainer.style.display = 'block';
                    }
                };

                // すべてのリスナーを削除して重複を防ぐ
                creationMethodRadios.forEach(radio => {
                    radio.removeEventListener('change', handleCreationMethodChange);
                });

                // 新しいリスナーを追加
                creationMethodRadios.forEach(radio => {
                    radio.addEventListener('change', handleCreationMethodChange);
                });

                // ラジオボタンを新規作成に初期化
                document.getElementById('creation-method-new').checked = true;
                document.getElementById('creation-method-copy').checked = false;

                const populateCopySourceSelect = () => {
                    copySourceSelect.innerHTML = '';
                    const defaultOption = document.createElement('option');
                    defaultOption.textContent = 'コピー元を選択...';
                    defaultOption.value = '';
                    copySourceSelect.appendChild(defaultOption);

                    state.financialData.forEach(year => {
                        const option = document.createElement('option');
                        option.value = year.id;
                        option.textContent = year.year;
                        copySourceSelect.appendChild(option);
                    });
                };

                // 初期表示（新規作成モードに設定）
                handleCreationMethodChange();

                showModal(addYearModal);
            };

            if (addYearForm) {
                addYearForm.addEventListener('submit', (e) => {
                    e.preventDefault();
                    const name = newYearNameInput.value.trim();
                    if (!name) {
                        showPopupMessage('年度/年 名前を入力してください。');
                        return;
                    }

                    // 「新規作成」か「コピーして作成」かを判定
                    const creationMethod = document.querySelector('input[name="creation-method"]:checked').value;

                    if (creationMethod === 'copy') {
                        // コピーして作成モード
                        const copySourceYearId = document.getElementById('copy-source-year').value;
                        if (!copySourceYearId) {
                            showPopupMessage('コピー元の年度/年を選択してください。');
                            return;
                        }

                        const sourceYear = findYear(copySourceYearId);
                        if (!sourceYear) {
                            showPopupMessage('コピー元の年度/年が見つかりません。');
                            return;
                        }

                        // 元の年度をディープコピー
                        const newYear = JSON.parse(JSON.stringify(sourceYear));
                        newYear.id = `year_${Date.now()}_${Math.random()}`;
                        newYear.year = name;

                        // 月と項目のIDを新しいものに振り直す
                        newYear.months.forEach(month => {
                            month.id = `month_${Date.now()}_${Math.random()}`;
                            month.income.forEach(item => {
                                item.id = `item_${Date.now()}_${Math.random()}`;
                            });
                            month.expenditure.forEach(item => {
                                item.id = `item_${Date.now()}_${Math.random()}`;
                            });
                        });

                        // 元の年度の後ろに追加
                        const sourceIndex = state.financialData.findIndex(y => y.id === copySourceYearId);
                        if (sourceIndex !== -1) {
                            state.financialData.splice(sourceIndex + 1, 0, newYear);
                        } else {
                            state.financialData.push(newYear);
                        }

                        saveState();
                        calculateAllBalances();
                        rerender();
                        showPopupMessage(`「${name}」が作成されました（コピー元: ${sourceYear.year}）。`);
                    } else {
                        // 新規作成モード
                        const afterYear = findYear(currentAfterYearId);
                        const isFiscalType = afterYear ? (afterYear.year.includes('年度') || afterYear.year.includes('年次')) : false;
                        const type = isFiscalType ? 'fiscal' : 'calendar';

                        const selectedMonths = Array.from(addYearMonthsContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);

                        if (selectedMonths.length === 0) {
                            showPopupMessage('少なくとも1つの月を選択してください。');
                            return;
                        }

                        handleAddYear({ name, type, selectedMonths, afterYearId: currentAfterYearId });
                    }

                    hideModal(addYearModal);
                });
            }

            if (addYearModal) {
                addYearCloseButton.addEventListener('click', () => hideModal(addYearModal));
                addYearCancelButton.addEventListener('click', () => hideModal(addYearModal));
                addYearModal.addEventListener('click', (e) => {
                    if (e.target === addYearModal) hideModal(addYearModal);
                });
            }

            // --- Edit Months Modal ---
            let currentEditingYearId = null;

            const openEditMonthsModal = (yearId) => {
                currentEditingYearId = yearId;
                const year = findYear(yearId);
                if (!year) return;

                editMonthsTitle.textContent = `「${year.year}」の月の編集`;

                const isFiscal = year.year.includes('年度') || year.year.includes('年次');
                const allMonths = isFiscal
                    ? ['4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月', '1月', '2月', '3月']
                    : ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];

                const existingMonthNames = new Set(year.months.map(m => m.month));

                editMonthsContainer.innerHTML = allMonths.map(monthName => {
                    const monthData = year.months.find(m => m.month === monthName);
                    const hasData = monthData && (monthData.income.length > 0 || monthData.expenditure.length > 0 || monthData.note);
                    const isChecked = existingMonthNames.has(monthName);

                    return `
                <div class="month-checkbox-item" title="${hasData ? 'データが含まれているため削除できません' : ''}">
                    <input type="checkbox" id="edit-month-${monthName}" value="${monthName}" ${isChecked ? 'checked' : ''} ${hasData ? 'disabled' : ''}>
                    <label for="edit-month-${monthName}" class="${hasData ? 'disabled' : ''}">${monthName}</label>
                </div>
            `;
                }).join('');

                showModal(editMonthsModal);
            };

            if (editMonthsForm) {
                editMonthsForm.addEventListener('submit', (e) => {
                    e.preventDefault();
                    const year = findYear(currentEditingYearId);
                    if (!year) return;

                    const selectedMonthNames = new Set(Array.from(editMonthsContainer.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value));
                    const existingMonthNames = new Set(year.months.map(m => m.month));

                    // Remove months
                    year.months = year.months.filter(month => selectedMonthNames.has(month.month));

                    // Add new months
                    selectedMonthNames.forEach(monthName => {
                        if (!existingMonthNames.has(monthName)) {
                            year.months.push({
                                id: `month_${Date.now()}_${Math.random()}`,
                                month: monthName,
                                note: '',
                                startingBalance: 0, // Will be recalculated
                                income: [],
                                expenditure: []
                            });
                        }
                    });

                    hideModal(editMonthsModal);
                    rerender();
                });
            }

            if (deleteYearButton) {
                deleteYearButton.addEventListener('click', () => {
                    if (!currentEditingYearId) return;
                    if (confirm('この年を完全に削除しますか？この操作は元に戻せません。')) {
                        state.financialData = state.financialData.filter(year => year.id !== currentEditingYearId);
                        hideModal(editMonthsModal);
                        rerender();
                    }
                });
            }

            if (editMonthsModal) {
                editMonthsCloseButton.addEventListener('click', () => hideModal(editMonthsModal));
                editMonthsCancelButton.addEventListener('click', () => hideModal(editMonthsModal));
                editMonthsModal.addEventListener('click', (e) => {
                    if (e.target === editMonthsModal) hideModal(editMonthsModal);
                });
            }

            // --- Bulk Edit ---
            const bulkEditMonthsContainer = document.getElementById('bulk-edit-months');

            function getBulkEditMonthCheckboxes() {
                if (!bulkEditMonthsContainer) return [];
                return Array.from(bulkEditMonthsContainer.querySelectorAll('input[type="checkbox"][data-month-id]'));
            }

            function syncBulkEditSelectAllStates() {
                if (!bulkEditMonthsContainer) return;
                const monthCheckboxes = getBulkEditMonthCheckboxes();
                const enabledMonths = monthCheckboxes.filter(cb => !cb.disabled);
                const master = document.getElementById('bulk-month-all');
                if (master) {
                    master.disabled = enabledMonths.length === 0;
                    master.parentElement?.classList.toggle('disabled', master.disabled);
                    master.checked = !master.disabled && enabledMonths.every(cb => cb.checked);
                }

                const yearGroups = bulkEditMonthsContainer.querySelectorAll('.month-checkbox-group[data-year-id]');
                yearGroups.forEach(group => {
                    const yearCheckbox = group.querySelector('.select-all-year input[type="checkbox"]');
                    if (!yearCheckbox) return;
                    const childMonths = Array.from(group.querySelectorAll('input[type="checkbox"][data-month-id]'));
                    const enabledChildren = childMonths.filter(cb => !cb.disabled);
                    yearCheckbox.disabled = enabledChildren.length === 0;
                    yearCheckbox.parentElement?.classList.toggle('disabled', yearCheckbox.disabled);
                    if (yearCheckbox.disabled) {
                        yearCheckbox.checked = false;
                    } else {
                        yearCheckbox.checked = enabledChildren.length > 0 && enabledChildren.every(cb => cb.checked);
                    }
                });
            }

            function updateBulkEditMonthAvailability() {
                if (!bulkEditMonthsContainer) return;
                const selectedName = bulkEditSelect.value;
                const monthCheckboxes = getBulkEditMonthCheckboxes();

                monthCheckboxes.forEach(cb => {
                    const wrapper = cb.parentElement;
                    if (!selectedName) {
                        cb.disabled = true;
                        cb.checked = false;
                        wrapper?.classList.add('disabled');
                        return;
                    }

                    const yearId = cb.getAttribute('data-year-id');
                    const monthId = cb.getAttribute('data-month-id');
                    const monthData = findMonth(yearId, monthId);
                    const hasItem = !!monthData && (
                        monthData.income.some(item => item.name === selectedName) ||
                        monthData.expenditure.some(item => item.name === selectedName)
                    );

                    cb.disabled = !hasItem;
                    if (!hasItem) cb.checked = false;
                    wrapper?.classList.toggle('disabled', !hasItem);
                });

                syncBulkEditSelectAllStates();
            }


            const populateBulkEditMonths = () => {
                let content = '';
                const selectAllGlobalLabel = '\u3059\u3079\u3066\u9078\u629c/\u89e3\u9664';
                const selectAllYearLabel = '\u3053\u306e\u5e74\u306e\u6708\u3092\u3059\u3079\u3066\u9078\u629c';
                const createCheckbox = (id, label, value, isChecked = true, customClass = '', dataAttrs = '') => `
            <div class=\"month-checkbox-item ${customClass}\">
                <input type=\"checkbox\" id=\"${id}\" value=\"${value}\" ${isChecked ? 'checked' : ''} ${dataAttrs}>
                <label for=\"${id}\">${label}</label>
            </div>`;

                content += `<div class=\"month-checkbox-group\">
            ${createCheckbox('bulk-month-all', selectAllGlobalLabel, 'all', true, 'select-all-global')}
        </div>`;

                state.financialData.forEach(year => {
                    content += `<div class=\"month-checkbox-group\" data-year-id=\"${year.id}\">`;
                    content += `<span class=\"month-checkbox-group-title\">${year.year}</span>`;
                    const yearCheckboxId = `bulk-month-year-${year.id}`;
                    content += createCheckbox(yearCheckboxId, selectAllYearLabel, year.id, true, 'select-all-year');

                    year.months.forEach(month => {
                        content += createCheckbox(`bulk-month-${month.id}`, month.month, month.id, true, '', `data-year-id=\"${year.id}\" data-month-id=\"${month.id}\"`);
                    });
                    content += `</div>`;
                });
                bulkEditMonthsContainer.innerHTML = content;
                updateBulkEditMonthAvailability();
            };

            const populateBulkEditSelect = () => {
                const allItems = extractCategoriesFromFinancialData();
                bulkEditSelect.innerHTML = '<option value="">編集する項目を選択...</option>';
                allItems.forEach(name => {
                    const option = document.createElement('option');
                    option.value = name;
                    option.textContent = name;
                    bulkEditSelect.appendChild(option);
                });
            };

            bulkEditButton.addEventListener('click', () => {
                try {
                    populateBulkEditSelect();
                    populateBulkEditMonths();
                } catch (err) {
                    console.error('Error preparing bulk edit modal:', err);
                }
                if (bulkEditModal) showModal(bulkEditModal);
            });

            bulkCancelButton.addEventListener('click', () => hideModal(bulkEditModal));
            bulkEditModal.addEventListener('click', (e) => {
                if (e.target === bulkEditModal) hideModal(bulkEditModal);
            });
            if (bulkAddModal) {
                bulkAddModal.addEventListener('click', (e) => {
                    if (e.target === bulkAddModal) hideModal(bulkAddModal);
                });
            }

            bulkEditSelect.addEventListener('change', updateBulkEditMonthAvailability);

            bulkEditMonthsContainer.addEventListener('change', (e) => {
                const target = e.target;
                if (target.type !== 'checkbox' || target.disabled) return;
                const isChecked = target.checked;

                if (target.value === 'all') {
                    bulkEditMonthsContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                        if (!cb.disabled) cb.checked = isChecked;
                    });
                } else if (target.parentElement.classList.contains('select-all-year')) {
                    const yearGroup = target.closest('.month-checkbox-group');
                    yearGroup.querySelectorAll('input[type="checkbox"]').forEach(cb => {
                        if (!cb.disabled) cb.checked = isChecked;
                    });
                }

                syncBulkEditSelectAllStates();
            });
            bulkEditForm.addEventListener('submit', e => {
                e.preventDefault();
                const oldName = bulkEditSelect.value;
                const newName = bulkEditNewName.value.trim();
                const newAmountStr = bulkEditNewAmount.value;
                const newAmount = newAmountStr ? parseInt(newAmountStr, 10) : null;

                const selectedMonthIds = Array.from(bulkEditMonthsContainer.querySelectorAll('input[type="checkbox"][data-month-id]:checked'))
                    .map(cb => cb.value);

                if (!oldName) return showPopupMessage('編集する項目を選択してください。');
                if (newName === '' && newAmount === null) return showPopupMessage('新しい項目名または新しい金額の少なくとも一方を入力してください。');
                if (newAmount !== null && isNaN(newAmount)) return showPopupMessage('有効な金額を入力してください。');
                if (selectedMonthIds.length === 0) return showPopupMessage('適用する月を少なくとも1つ選択してください。');

                let updatedCount = 0;
                state.financialData.forEach(year => {
                    year.months.forEach(month => {
                        if (selectedMonthIds.includes(month.id)) {
                            const updateItems = (items) => {
                                items.forEach(item => {
                                    if (item.name === oldName) {
                                        if (newName !== '') item.name = newName;
                                        if (newAmount !== null) item.amount = newAmount;
                                        updatedCount++;
                                    }
                                });
                            };
                            updateItems(month.income);
                            updateItems(month.expenditure);
                        }
                    });
                });

                if (updatedCount > 0) {
                    if (confirm(`${updatedCount}個の項目を更新します。よろしいですか？`)) {
                        bulkEditNewName.value = '';
                        bulkEditNewAmount.value = '';
                        hideModal(bulkEditModal);
                        rerender();
                    }
                } else {
                    showPopupMessage('選択された月に該当する項目が見つかりませんでした。');
                    hideModal(bulkEditModal);
                }
            });

            // Populate the select/datalist for bulk add item names
            const populateBulkAddSelect = () => {
                const allItems = new Set();
                state.financialData.forEach(year => {
                    year.months.forEach(month => {
                        month.income.forEach(item => allItems.add(item.name));
                        month.expenditure.forEach(item => allItems.add(item.name));
                    });
                });

                const bulkAddNameInput = document.getElementById('bulk-add-name');
                const datalistId = bulkAddNameInput.getAttribute('list');
                const datalist = document.getElementById(datalistId);
                if (datalist) datalist.innerHTML = ''; // Clear existing options

                const sortedItems = Array.from(allItems).sort();
                sortedItems.forEach(name => {
                    const option = document.createElement('option');
                    option.value = name;
                    if (datalist) datalist.appendChild(option);
                });
            };


            function showSummaryModal() {
                const modalContent = createSummaryContent();
                openModal('収支概要', modalContent, 'wide');
            }

            function createSummaryContent() {
                const totalIncome = Object.values(appData.years).reduce((sum, year) => sum + year.income, 0);
                const totalExpenditure = Object.values(appData.years).reduce((sum, year) => sum + year.expenditure, 0);
                const netTotal = totalIncome - totalExpenditure;

                return `
            <div class="summary-overview">
                <div class="summary-item">
                    <span class="summary-label">総収入:</span>
                    <span class="summary-value income">${formatCurrency(totalIncome)}</span>
                </div>
                <div class="summary-item">
                    <span class="summary-label">総支出:</span>
                    <span class="summary-value expenditure">${formatCurrency(totalExpenditure)}</span>
                </div>
                <div class="summary-item">
                    <span class="summary-label">ネット合計:</span>
                    <span class="summary-value net">${formatCurrency(netTotal)}</span>
                </div>
            </div>
            <h3>年別詳細</h3>
            <div class="summary-year-details">
                ${Object.keys(appData.years).map(year => {
                    const yearData = appData.years[year];
                    const yearTotalIncome = yearData.income;
                    const yearTotalExpenditure = yearData.expenditure;
                    const yearNetTotal = yearTotalIncome - yearTotalExpenditure;

                    return `
                        <div class="summary-year-item">
                            <div class="summary-year-header">
                                <span class="summary-year-title">${year}年</span>
                                <span class="summary-year-net ${yearNetTotal >= 0 ? 'positive' : 'negative'}">${formatCurrency(yearNetTotal)}</span>
                            </div>
                            <div class="summary-year-content">
                                <div class="summary-item">
                                    <span class="summary-label">収入:</span>
                                    <span class="summary-value income">${formatCurrency(yearTotalIncome)}</span>
                                </div>
                                <div class="summary-item">
                                    <span class="summary-label">支出:</span>
                                    <span class="summary-value expenditure">${formatCurrency(yearTotalExpenditure)}</span>
                                </div>
                            </div>
                        </div>
                    `;
                }).join('')}
            </div>
        `;
            }

            function formatCurrency(amount) {
                return new Intl.NumberFormat('ja-JP', { style: 'currency', currency: 'JPY' }).format(amount);
            }

            // Initialize the app
            initializeApp();
        });

        // ポップアップメッセージを表示する関数
        function showPopupMessage(message) {
            const popupMessage = document.getElementById('popupMessage');
            const popupText = document.getElementById('popupText');
            const closePopup = document.getElementById('closePopup');

            popupText.textContent = message;
            popupMessage.style.display = 'flex'; // ポップアップを表示

            closePopup.onclick = () => {
                popupMessage.style.display = 'none'; // ポップアップを非表示
            };

            // ポップアップの外側をクリックしても閉じるようにする
            popupMessage.onclick = (event) => {
                if (event.target === popupMessage) {
                    popupMessage.style.display = 'none';
                }
            };
        }

    