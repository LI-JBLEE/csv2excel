// CSV to Excel Converter Application with Mapping Management

class CSVToExcelConverter {
    constructor() {
        this.csvData = null;
        this.convertedData = null;
        this.currentTab = 'converter';
        this.init();
    }

    init() {
        // Wait for DOM to be fully loaded
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', () => {
                this.initializeElements();
                this.attachEventListeners();
                this.setupMappings();
                this.renderMappingTables();
            });
        } else {
            this.initializeElements();
            this.attachEventListeners();
            this.setupMappings();
            this.renderMappingTables();
        }
    }

    initializeElements() {
        // Tab elements
        this.tabBtns = document.querySelectorAll('.tab-btn');
        this.tabContents = document.querySelectorAll('.tab-content');

        // Converter tab elements
        this.uploadArea = document.getElementById('uploadArea');
        this.fileInput = document.getElementById('fileInput');
        this.fileInfo = document.getElementById('fileInfo');
        this.fileName = document.getElementById('fileName');
        this.fileSize = document.getElementById('fileSize');
        this.removeFileBtn = document.getElementById('removeFile');
        this.convertBtn = document.getElementById('convertBtn');
        this.convertBtnText = document.getElementById('convertBtnText');
        this.convertSpinner = document.getElementById('convertSpinner');
        this.statusSection = document.getElementById('statusSection');
        this.statusMessage = document.getElementById('statusMessage');
        this.previewSection = document.getElementById('previewSection');
        this.previewTable = document.getElementById('previewTable');
        this.previewTableBody = document.getElementById('previewTableBody');
        this.previewCount = document.getElementById('previewCount');
        this.downloadSection = document.getElementById('downloadSection');
        this.downloadBtn = document.getElementById('downloadBtn');

        // Mapping tab elements
        this.countryMappingTableBody = document.getElementById('countryMappingTableBody');
        this.payCodeMappingTableBody = document.getElementById('payCodeMappingTableBody');
        this.countryJsonEditor = document.getElementById('countryJsonEditor');
        this.payCodeJsonEditor = document.getElementById('payCodeJsonEditor');
        this.addCountryBtn = document.getElementById('addCountryBtn');
        this.addPayCodeBtn = document.getElementById('addPayCodeBtn');
        this.updateCountryJsonBtn = document.getElementById('updateCountryJsonBtn');
        this.updatePayCodeJsonBtn = document.getElementById('updatePayCodeJsonBtn');
        this.resetCountryMappingsBtn = document.getElementById('resetCountryMappingsBtn');
        this.resetPayCodeMappingsBtn = document.getElementById('resetPayCodeMappingsBtn');

        console.log('Elements initialized');
    }

    setupMappings() {
        // Default mappings from the provided data
        this.defaultPayCodeMapping = {
            "Bonus-460": "DA",
            "Bonus-500": "DD",
            "Comm-300": "CV",
            "Draws-305": "DN",
            "Draws-310": "DO",
            "OKR-450": "OM",
            "SPIFF-430": "S3"
        };

        this.defaultCountryMapping = {
            "Australia": { code: "AU", currency: "AUD" },
            "Singapore": { code: "SG", currency: "SGD" },
            "Malaysia": { code: "MY", currency: "MYR" },
            "Thailand": { code: "TH", currency: "THB" },
            "Philippines": { code: "PH", currency: "PHP" },
            "Indonesia": { code: "ID", currency: "IDR" },
            "Vietnam": { code: "VN", currency: "VND" },
            "Japan": { code: "JP", currency: "JPY" },
            "Korea": { code: "KR", currency: "KRW" }
        };

        // Current mappings (can be modified by user)
        this.payCodeMapping = { ...this.defaultPayCodeMapping };
        this.countryMapping = { ...this.defaultCountryMapping };

        // Fixed values
        this.effectiveDate = "2025-06-01";
    }

    attachEventListeners() {
        console.log('Attaching event listeners...');

        // Tab switching
        this.tabBtns.forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                this.switchTab(btn.dataset.tab);
            });
        });

        // File upload events
        if (this.uploadArea && this.fileInput) {
            this.uploadArea.addEventListener('click', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.fileInput.click();
            });
            
            this.fileInput.addEventListener('change', (e) => {
                this.handleFileSelect(e);
            });
        }
        
        // Drag and drop events
        if (this.uploadArea) {
            this.uploadArea.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.handleDragOver(e);
            });
            
            this.uploadArea.addEventListener('dragleave', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.handleDragLeave(e);
            });
            
            this.uploadArea.addEventListener('drop', (e) => {
                e.preventDefault();
                e.stopPropagation();
                this.handleDrop(e);
            });
        }

        // Prevent default drag behaviors on document
        document.addEventListener('dragover', (e) => e.preventDefault());
        document.addEventListener('drop', (e) => e.preventDefault());

        // Converter events
        if (this.removeFileBtn) {
            this.removeFileBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.removeFile();
            });
        }
        
        if (this.convertBtn) {
            this.convertBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.convertData();
            });
        }
        
        if (this.downloadBtn) {
            this.downloadBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.downloadExcel();
            });
        }

        // Mapping management events
        if (this.addCountryBtn) {
            this.addCountryBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.addCountryMapping();
            });
        }

        if (this.addPayCodeBtn) {
            this.addPayCodeBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.addPayCodeMapping();
            });
        }

        if (this.updateCountryJsonBtn) {
            this.updateCountryJsonBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.updateCountryMappingsFromJson();
            });
        }

        if (this.updatePayCodeJsonBtn) {
            this.updatePayCodeJsonBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.updatePayCodeMappingsFromJson();
            });
        }

        if (this.resetCountryMappingsBtn) {
            this.resetCountryMappingsBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.resetCountryMappings();
            });
        }

        if (this.resetPayCodeMappingsBtn) {
            this.resetPayCodeMappingsBtn.addEventListener('click', (e) => {
                e.preventDefault();
                this.resetPayCodeMappings();
            });
        }

        console.log('Event listeners attached');
    }

    switchTab(tabName) {
        this.currentTab = tabName;
        
        // Update tab buttons
        this.tabBtns.forEach(btn => {
            btn.classList.toggle('active', btn.dataset.tab === tabName);
        });

        // Update tab content
        this.tabContents.forEach(content => {
            content.classList.toggle('active', content.id === `${tabName}-tab`);
        });

        // Update JSON editors when switching to mappings tab
        if (tabName === 'mappings') {
            this.updateJsonEditors();
        }
    }

    // File handling methods
    handleDragOver(e) {
        this.uploadArea.classList.add('drag-over');
    }

    handleDragLeave(e) {
        if (!this.uploadArea.contains(e.relatedTarget)) {
            this.uploadArea.classList.remove('drag-over');
        }
    }

    handleDrop(e) {
        this.uploadArea.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            this.processFile(files[0]);
        }
    }

    handleFileSelect(e) {
        const file = e.target.files[0];
        if (file) {
            this.processFile(file);
        }
    }

    processFile(file) {
        // Validate file type
        if (!file.name.toLowerCase().endsWith('.csv')) {
            this.showStatus('Please select a CSV file.', 'error');
            return;
        }

        // Update UI to show file info
        this.fileName.textContent = file.name;
        this.fileSize.textContent = this.formatFileSize(file.size);
        this.fileInfo.style.display = 'flex';
        this.uploadArea.style.display = 'none';
        
        // Enable convert button
        this.convertBtn.disabled = false;
        this.convertBtn.classList.remove('btn--disabled');

        // Read file content
        const reader = new FileReader();
        reader.onload = (e) => {
            this.csvData = e.target.result;
            this.showStatus('CSV file loaded successfully. Click "Convert to Excel" to process the data.', 'success');
        };
        
        reader.onerror = (e) => {
            this.showStatus('Error reading the file. Please try again.', 'error');
            this.removeFile();
        };
        
        reader.readAsText(file);
    }

    removeFile() {
        // Reset file input
        this.fileInput.value = '';
        this.csvData = null;
        this.convertedData = null;
        
        // Reset UI
        this.fileInfo.style.display = 'none';
        this.uploadArea.style.display = 'block';
        this.convertBtn.disabled = true;
        this.convertBtn.classList.add('btn--disabled');
        this.previewSection.style.display = 'none';
        this.downloadSection.style.display = 'none';
        
        this.hideStatus();
    }

    formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    parseCSV(csvText) {
        const lines = csvText.trim().split('\n');
        if (lines.length < 2) {
            throw new Error('CSV file must contain at least a header row and one data row.');
        }
        
        const headers = this.parseCSVLine(lines[0]);
        const data = [];

        for (let i = 1; i < lines.length; i++) {
            const line = lines[i].trim();
            if (!line) continue;
            
            const values = this.parseCSVLine(line);
            if (values.length > 0) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header.trim()] = values[index] ? values[index].trim() : '';
                });
                data.push(row);
            }
        }

        return { headers, data };
    }

    parseCSVLine(line) {
        const result = [];
        let current = '';
        let inQuotes = false;
        let i = 0;

        while (i < line.length) {
            const char = line[i];
            
            if (char === '"') {
                if (inQuotes && line[i + 1] === '"') {
                    current += '"';
                    i += 2;
                } else {
                    inQuotes = !inQuotes;
                    i++;
                }
            } else if (char === ',' && !inQuotes) {
                result.push(current);
                current = '';
                i++;
            } else {
                current += char;
                i++;
            }
        }
        
        result.push(current);
        return result;
    }

    async convertData() {
        if (!this.csvData) {
            this.showStatus('No CSV data to convert.', 'error');
            return;
        }

        this.setLoading(true);
        this.showStatus('Converting data...', 'info');

        try {
            await new Promise(resolve => setTimeout(resolve, 500));
            
            const { headers, data } = this.parseCSV(this.csvData);
            const transformedData = this.transformData(data, headers);
            
            if (transformedData.length === 0) {
                throw new Error('No valid data found to convert. Please check your CSV file format and ensure it contains the required columns: Employee Number, Country, Business Unit, and Pay Code columns with (LCY) suffix.');
            }

            this.convertedData = transformedData;
            this.showPreview(transformedData);
            this.downloadSection.style.display = 'block';
            this.showStatus(`Successfully converted ${transformedData.length} records.`, 'success');

        } catch (error) {
            console.error('Conversion error:', error);
            this.showStatus(`Error: ${error.message}`, 'error');
        } finally {
            this.setLoading(false);
        }
    }

    transformData(data, headers) {
        const transformedRows = [];

        // Find pay code columns
        const payCodeColumns = headers.filter(header => {
            return header.includes('(LCY)') && Object.keys(this.payCodeMapping).some(payCode => 
                header.includes(payCode)
            );
        });

        data.forEach((row, index) => {
            const employeeNumber = row['Employee Number'] || row['EmployeeNumber'] || row['Employee_Number'] || '';
            const country = row['Country'] || '';
            // FIX: Read Business Unit from the original CSV file
            const businessUnit = row['Business Unit'] || row['BusinessUnit'] || row['Business_Unit'] || '';

            if (!employeeNumber || !country) {
                console.warn(`Row ${index + 1}: Missing required data - Employee Number: ${employeeNumber}, Country: ${country}`);
                return;
            }

            const countryInfo = this.countryMapping[country];
            if (!countryInfo) {
                console.warn(`Row ${index + 1}: Unknown country: ${country}`);
                return;
            }

            payCodeColumns.forEach(column => {
                const amountStr = row[column] || '';
                if (!amountStr || amountStr.trim() === '' || amountStr === '0') {
                    return;
                }

                const amount = this.parseAmount(amountStr);
                if (amount === 0) {
                    return;
                }

                const payCodeMatch = Object.keys(this.payCodeMapping).find(payCode => 
                    column.includes(payCode)
                );

                if (!payCodeMatch) {
                    console.warn(`Unknown pay code in column: ${column}`);
                    return;
                }

                const mappedPayCode = this.payCodeMapping[payCodeMatch];
                const payCodeDescription = payCodeMatch;

                const transformedRow = {
                    'Country': countryInfo.code,
                    'Employee Number': employeeNumber,
                    'Paygroup': `${country}_Monthly`,
                    'Effective Date': this.effectiveDate,
                    'Type of Input': 'Payment',
                    'Pay Code': mappedPayCode,
                    'Pay Code description': payCodeDescription,
                    'Amount': amount,
                    'Currency': countryInfo.currency,
                    'Unit': 'A',
                    'Region': 'APAC',
                    'Business Unit': businessUnit // Use the actual value from CSV
                };

                transformedRows.push(transformedRow);
            });
        });

        return transformedRows;
    }

    parseAmount(amountStr) {
        const cleanAmount = amountStr.replace(/[",\s]/g, '');
        const parsed = parseFloat(cleanAmount);
        return isNaN(parsed) ? 0 : parsed;
    }

    showPreview(data) {
        this.previewTableBody.innerHTML = '';
        
        const previewData = data.slice(0, 10);
        
        previewData.forEach(row => {
            const tr = document.createElement('tr');
            
            const columns = [
                'Country', 'Employee Number', 'Paygroup', 'Effective Date',
                'Type of Input', 'Pay Code', 'Pay Code description', 'Amount',
                'Currency', 'Unit', 'Region', 'Business Unit'
            ];

            columns.forEach(column => {
                const td = document.createElement('td');
                td.textContent = row[column] || '';
                tr.appendChild(td);
            });

            this.previewTableBody.appendChild(tr);
        });

        this.previewCount.textContent = `Showing ${previewData.length} of ${data.length} records`;
        this.previewSection.style.display = 'block';
    }

    downloadExcel() {
        if (!this.convertedData || this.convertedData.length === 0) {
            this.showStatus('No data to download.', 'error');
            return;
        }

        try {
            const wb = XLSX.utils.book_new();
            
            const ws = XLSX.utils.json_to_sheet(this.convertedData, {
                header: [
                    'Country', 'Employee Number', 'Paygroup', 'Effective Date',
                    'Type of Input', 'Pay Code', 'Pay Code description', 'Amount',
                    'Currency', 'Unit', 'Region', 'Business Unit'
                ]
            });

            XLSX.utils.book_append_sheet(wb, ws, 'Payroll Data');

            const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
            
            const blob = new Blob([wbout], { type: 'application/octet-stream' });
            const url = URL.createObjectURL(blob);
            
            const a = document.createElement('a');
            a.href = url;
            a.download = `converted_payroll.xlsx`;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            this.showStatus('Excel file downloaded successfully!', 'success');

        } catch (error) {
            console.error('Download error:', error);
            this.showStatus(`Error downloading file: ${error.message}`, 'error');
        }
    }

    // Mapping management methods
    renderMappingTables() {
        this.renderCountryMappingTable();
        this.renderPayCodeMappingTable();
        this.updateJsonEditors();
    }

    renderCountryMappingTable() {
        this.countryMappingTableBody.innerHTML = '';
        
        Object.entries(this.countryMapping).forEach(([countryName, info]) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="text" value="${this.escapeHtml(countryName)}" data-field="name" data-original="${this.escapeHtml(countryName)}"></td>
                <td><input type="text" value="${this.escapeHtml(info.code)}" data-field="code" data-original="${this.escapeHtml(countryName)}"></td>
                <td><input type="text" value="${this.escapeHtml(info.currency)}" data-field="currency" data-original="${this.escapeHtml(countryName)}"></td>
                <td class="mapping-actions">
                    <button class="btn btn--xs btn--danger delete-country-btn" data-country="${this.escapeHtml(countryName)}">Delete</button>
                </td>
            `;
            
            // Add change listeners with better input handling
            tr.querySelectorAll('input').forEach(input => {
                input.addEventListener('focus', (e) => {
                    e.target.select(); // Select all text on focus for easier editing
                });
                
                input.addEventListener('blur', (e) => {
                    this.updateCountryMapping(e.target.dataset.original, e.target.dataset.field, e.target.value);
                });
                
                input.addEventListener('keypress', (e) => {
                    if (e.key === 'Enter') {
                        e.target.blur(); // Trigger blur event to save changes
                    }
                });
            });
            
            this.countryMappingTableBody.appendChild(tr);
        });

        // Add event listeners for delete buttons
        this.countryMappingTableBody.querySelectorAll('.delete-country-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                this.deleteCountryMapping(btn.dataset.country);
            });
        });
    }

    renderPayCodeMappingTable() {
        this.payCodeMappingTableBody.innerHTML = '';
        
        Object.entries(this.payCodeMapping).forEach(([originalCode, eyCode]) => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><input type="text" value="${this.escapeHtml(originalCode)}" data-field="original" data-original="${this.escapeHtml(originalCode)}"></td>
                <td><input type="text" value="${this.escapeHtml(eyCode)}" data-field="ey" data-original="${this.escapeHtml(originalCode)}"></td>
                <td class="mapping-actions">
                    <button class="btn btn--xs btn--danger delete-paycode-btn" data-paycode="${this.escapeHtml(originalCode)}">Delete</button>
                </td>
            `;
            
            // Add change listeners with better input handling
            tr.querySelectorAll('input').forEach(input => {
                input.addEventListener('focus', (e) => {
                    e.target.select(); // Select all text on focus for easier editing
                });
                
                input.addEventListener('blur', (e) => {
                    this.updatePayCodeMapping(e.target.dataset.original, e.target.dataset.field, e.target.value);
                });
                
                input.addEventListener('keypress', (e) => {
                    if (e.key === 'Enter') {
                        e.target.blur(); // Trigger blur event to save changes
                    }
                });
            });
            
            this.payCodeMappingTableBody.appendChild(tr);
        });

        // Add event listeners for delete buttons
        this.payCodeMappingTableBody.querySelectorAll('.delete-paycode-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                e.preventDefault();
                this.deletePayCodeMapping(btn.dataset.paycode);
            });
        });
    }

    addCountryMapping() {
        const timestamp = new Date().toLocaleTimeString('en-US', { hour12: false }).replace(/:/g, '');
        const newCountryName = `NewCountry_${timestamp}`;
        this.countryMapping[newCountryName] = { code: 'XX', currency: 'XXX' };
        this.renderCountryMappingTable();
        this.updateJsonEditors();
        this.showStatus('New country mapping added. Please edit the values.', 'info');
    }

    addPayCodeMapping() {
        const timestamp = new Date().toLocaleTimeString('en-US', { hour12: false }).replace(/:/g, '');
        const newPayCode = `NewPayCode-${timestamp}`;
        this.payCodeMapping[newPayCode] = 'XX';
        this.renderPayCodeMappingTable();
        this.updateJsonEditors();
        this.showStatus('New pay code mapping added. Please edit the values.', 'info');
    }

    updateCountryMapping(originalName, field, value) {
        if (!value.trim()) {
            this.showStatus('Value cannot be empty.', 'error');
            this.renderCountryMappingTable(); // Reset to original values
            return;
        }

        if (field === 'name' && value !== originalName) {
            if (this.countryMapping[value]) {
                this.showStatus('Country name already exists.', 'error');
                this.renderCountryMappingTable(); // Reset to original values
                return;
            }
            // Rename the mapping
            this.countryMapping[value] = this.countryMapping[originalName];
            delete this.countryMapping[originalName];
            this.renderCountryMappingTable();
        } else if (field === 'code') {
            this.countryMapping[originalName].code = value.toUpperCase();
        } else if (field === 'currency') {
            this.countryMapping[originalName].currency = value.toUpperCase();
        }
        this.updateJsonEditors();
    }

    updatePayCodeMapping(originalCode, field, value) {
        if (!value.trim()) {
            this.showStatus('Value cannot be empty.', 'error');
            this.renderPayCodeMappingTable(); // Reset to original values
            return;
        }

        if (field === 'original' && value !== originalCode) {
            if (this.payCodeMapping[value]) {
                this.showStatus('Pay code already exists.', 'error');
                this.renderPayCodeMappingTable(); // Reset to original values
                return;
            }
            // Rename the mapping
            this.payCodeMapping[value] = this.payCodeMapping[originalCode];
            delete this.payCodeMapping[originalCode];
            this.renderPayCodeMappingTable();
        } else if (field === 'ey') {
            this.payCodeMapping[originalCode] = value.toUpperCase();
        }
        this.updateJsonEditors();
    }

    deleteCountryMapping(countryName) {
        if (confirm(`Are you sure you want to delete the mapping for "${countryName}"?`)) {
            delete this.countryMapping[countryName];
            this.renderCountryMappingTable();
            this.updateJsonEditors();
            this.showStatus(`Country mapping for "${countryName}" deleted.`, 'success');
        }
    }

    deletePayCodeMapping(payCode) {
        if (confirm(`Are you sure you want to delete the mapping for "${payCode}"?`)) {
            delete this.payCodeMapping[payCode];
            this.renderPayCodeMappingTable();
            this.updateJsonEditors();
            this.showStatus(`Pay code mapping for "${payCode}" deleted.`, 'success');
        }
    }

    updateJsonEditors() {
        if (this.countryJsonEditor) {
            this.countryJsonEditor.value = JSON.stringify(this.countryMapping, null, 2);
        }
        if (this.payCodeJsonEditor) {
            this.payCodeJsonEditor.value = JSON.stringify(this.payCodeMapping, null, 2);
        }
    }

    updateCountryMappingsFromJson() {
        try {
            const newMappings = JSON.parse(this.countryJsonEditor.value);
            
            // Validate JSON structure
            for (const [country, info] of Object.entries(newMappings)) {
                if (!info || typeof info !== 'object' || !info.code || !info.currency) {
                    throw new Error(`Invalid structure for country "${country}". Expected format: {"code": "XX", "currency": "XXX"}`);
                }
            }
            
            this.countryMapping = newMappings;
            this.renderCountryMappingTable();
            this.showStatus('Country mappings updated successfully from JSON.', 'success');
        } catch (error) {
            this.showStatus(`JSON Error: ${error.message}`, 'error');
        }
    }

    updatePayCodeMappingsFromJson() {
        try {
            const newMappings = JSON.parse(this.payCodeJsonEditor.value);
            
            // Validate JSON structure
            for (const [payCode, eyCode] of Object.entries(newMappings)) {
                if (!eyCode || typeof eyCode !== 'string') {
                    throw new Error(`Invalid structure for pay code "${payCode}". Expected string value.`);
                }
            }
            
            this.payCodeMapping = newMappings;
            this.renderPayCodeMappingTable();
            this.showStatus('Pay code mappings updated successfully from JSON.', 'success');
        } catch (error) {
            this.showStatus(`JSON Error: ${error.message}`, 'error');
        }
    }

    resetCountryMappings() {
        if (confirm('Are you sure you want to reset all country mappings to defaults? This will lose any custom changes.')) {
            this.countryMapping = { ...this.defaultCountryMapping };
            this.renderCountryMappingTable();
            this.updateJsonEditors();
            this.showStatus('Country mappings reset to defaults.', 'success');
        }
    }

    resetPayCodeMappings() {
        if (confirm('Are you sure you want to reset all pay code mappings to defaults? This will lose any custom changes.')) {
            this.payCodeMapping = { ...this.defaultPayCodeMapping };
            this.renderPayCodeMappingTable();
            this.updateJsonEditors();
            this.showStatus('Pay code mappings reset to defaults.', 'success');
        }
    }

    escapeHtml(text) {
        const div = document.createElement('div');
        div.textContent = text;
        return div.innerHTML;
    }

    setLoading(isLoading) {
        this.convertBtn.disabled = isLoading;
        if (isLoading) {
            this.convertBtn.classList.add('btn--disabled');
        } else {
            this.convertBtn.classList.remove('btn--disabled');
        }
        this.convertBtnText.textContent = isLoading ? 'Converting...' : 'Convert to Excel';
        this.convertSpinner.style.display = isLoading ? 'inline-block' : 'none';
    }

    showStatus(message, type = 'info') {
        this.statusMessage.textContent = message;
        this.statusMessage.className = `status-message ${type}`;
        this.statusSection.style.display = 'block';
        
        if (type === 'success' || type === 'info') {
            setTimeout(() => this.hideStatus(), 5000);
        }
    }

    hideStatus() {
        this.statusSection.style.display = 'none';
    }
}

// Initialize the application and make it globally accessible
console.log('Script loaded, initializing converter...');
const app = new CSVToExcelConverter();