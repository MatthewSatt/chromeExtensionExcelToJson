// popup.js
document.addEventListener('DOMContentLoaded', function() {
    // Get DOM elements
    const fileInput = document.getElementById('fileInput');
    const fileName = document.getElementById('fileName');
    const sheetSelect = document.getElementById('sheetSelect');
    const headerRowCheck = document.getElementById('headerRowCheck');
    const prettifyCheck = document.getElementById('prettifyCheck');
    const statusMessage = document.getElementById('statusMessage');
    const convertBtn = document.getElementById('convertBtn');
    const cancelBtn = document.getElementById('cancelBtn');
    const fileInputLabel = document.querySelector('.file-input-label');
    const resultsContainer = document.getElementById('resultsContainer');
    const jsonOutput = document.getElementById('jsonOutput');
    const copyJsonBtn = document.getElementById('copyJsonBtn');
    const spinner = document.getElementById('spinner');

    // Variable to store the latest JSON string
    let currentJsonString = '';

    // Function to highlight JSON syntax
    function syntaxHighlight(json) {
        json = json.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
        return json.replace(/("(\\u[a-zA-Z0-9]{4}|\\[^u]|[^\\"])*"(\s*:)?|\b(true|false|null)\b|-?\d+(?:\.\d*)?(?:[eE][+\-]?\d+)?)/g, function (match) {
            let cls = 'json-number';
            if (/^"/.test(match)) {
                if (/:$/.test(match)) {
                    cls = 'json-key';
                } else {
                    cls = 'json-string';
                }
            } else if (/true|false/.test(match)) {
                cls = 'json-boolean';
            } else if (/null/.test(match)) {
                cls = 'json-null';
            }
            return '<span class="' + cls + '">' + match + '</span>';
        });
    }

    // Function to show status message
    function showStatus(message, isSuccess = true) {
        statusMessage.textContent = message;
        statusMessage.className = 'status-message';
        statusMessage.classList.add(isSuccess ? 'status-success' : 'status-error');
        statusMessage.style.display = 'block';
    }

    // Function to hide status message
    function hideStatus() {
        statusMessage.style.display = 'none';
    }

    // Function to show loading spinner
    function showSpinner() {
        spinner.style.display = 'block';
        convertBtn.disabled = true;
    }

    // Function to hide loading spinner
    function hideSpinner() {
        spinner.style.display = 'none';
        convertBtn.disabled = false;
    }

    // Function to read an Excel file and return the sheet names
    function getExcelSheets(file) {
        return new Promise((resolve, reject) => {
            showSpinner();

            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    hideSpinner();
                    resolve(workbook.SheetNames);
                } catch (error) {
                    hideSpinner();
                    reject(error);
                }
            };

            reader.onerror = function(error) {
                hideSpinner();
                reject(error);
            };

            reader.readAsArrayBuffer(file);
        });
    }

    // Function to convert an Excel sheet to JSON
    function excelToJson(file, sheetName, hasHeader) {
        return new Promise((resolve, reject) => {
            showSpinner();

            const reader = new FileReader();

            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });

                    // Get the specified worksheet
                    const worksheet = workbook.Sheets[sheetName];

                    // Convert to JSON
                    const jsonOptions = {
                        header: hasHeader ? 1 : undefined,
                        defval: '',
                        blankrows: false
                    };

                    const jsonData = XLSX.utils.sheet_to_json(worksheet, jsonOptions);
                    hideSpinner();
                    resolve(jsonData);
                } catch (error) {
                    hideSpinner();
                    reject(error);
                }
            };

            reader.onerror = function(error) {
                hideSpinner();
                reject(error);
            };

            reader.readAsArrayBuffer(file);
        });
    }

    // Function to create a typing effect for the JSON output
    function typeJsonOutput(jsonString, prettify) {
        // Store the current JSON string for copying
        currentJsonString = jsonString;

        // Prepare the JSON with syntax highlighting if prettified
        const formattedJson = prettify ? syntaxHighlight(jsonString) : jsonString.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

        // Make sure the container is visible
        resultsContainer.style.display = 'block';

        // Clear previous content
        jsonOutput.innerHTML = '';

        // Type the JSON with a fast typing effect
        let i = 0;
        const speed = 1; // milliseconds per character (very fast)
        const chunkSize = 100; // characters per chunk

        function typeChunk() {
            if (i < formattedJson.length) {
                const end = Math.min(i + chunkSize, formattedJson.length);
                const chunk = formattedJson.substring(i, end);

                // For syntax highlighted JSON, we need to use innerHTML
                // For plain JSON, we can use textContent which is faster
                if (prettify) {
                    // Create a temporary div to parse the HTML
                    const tempDiv = document.createElement('div');
                    tempDiv.innerHTML = chunk;

                    // Append each child node to jsonOutput
                    while (tempDiv.firstChild) {
                        jsonOutput.appendChild(tempDiv.firstChild);
                    }
                } else {
                    jsonOutput.textContent += chunk;
                }

                jsonOutput.scrollTop = jsonOutput.scrollHeight;
                i = end;
                setTimeout(typeChunk, speed);
            }
        }

        typeChunk();
    }

    // CRITICAL FIX: Direct file input approach that works with Chrome extension popups
    // This approach prevents the popup from closing when selecting a file
    fileInput.addEventListener('click', function(e) {
        // Prevent the default behavior to avoid popup closure
        e.preventDefault();
        e.stopPropagation();

        // Create a communication channel to the background script
        // This helps keep the popup open during file selection
        chrome.runtime.sendMessage({action: "keepPopupOpen"});

        // Use a timeout to allow the message to be processed
        setTimeout(() => {
            // Programmatically open the file dialog without triggering popup closure
            const event = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window
            });

            // We need to modify the input to prevent default browser behavior
            // that causes popup closure
            const actualFileInput = document.createElement('input');
            actualFileInput.type = 'file';
            actualFileInput.accept = '.xlsx, .xls, .csv';
            actualFileInput.style.position = 'fixed';
            actualFileInput.style.opacity = '0';
            actualFileInput.style.pointerEvents = 'none';
            document.body.appendChild(actualFileInput);

            // Handle the file selection
            actualFileInput.addEventListener('change', async function() {
                if (actualFileInput.files && actualFileInput.files[0]) {
                    const selectedFile = actualFileInput.files[0];

                    // Update display and handle the selected file
                    fileName.textContent = selectedFile.name;
                    fileName.style.display = 'block';
                    hideStatus();
                    resultsContainer.style.display = 'none';

                    try {
                        // Get sheet names from the Excel file
                        const sheets = await getExcelSheets(selectedFile);
                        populateSheetSelect(sheets);

                        // Store the file data for later use
                        const dataTransfer = new DataTransfer();
                        dataTransfer.items.add(selectedFile);
                        fileInput.files = dataTransfer.files;
                    } catch (error) {
                        console.error('Error reading Excel file:', error);
                        showStatus('Error reading Excel file. Please make sure it\'s a valid Excel file.', false);
                        fileName.textContent = '';
                        resetSheetSelect();
                    }
                }

                // Clean up
                document.body.removeChild(actualFileInput);
            });

            // Trigger the file dialog
            actualFileInput.dispatchEvent(event);
        }, 50);
    });

    // Also update the label click handler to use the same approach
    fileInputLabel.addEventListener('click', function(e) {
        e.preventDefault();
        e.stopPropagation();

        // Simulate a click on the file input instead
        fileInput.click();
    });

    // Convert button click handler
    convertBtn.addEventListener('click', async function() {
        if (!fileInput.files || !fileInput.files[0]) {
            showStatus('Please select an Excel file first.', false);
            return;
        }

        const selectedFile = fileInput.files[0];
        const selectedSheet = sheetSelect.value;
        const hasHeader = headerRowCheck.checked;
        const prettify = prettifyCheck.checked;

        if (!selectedSheet) {
            showStatus('Please select a worksheet.', false);
            return;
        }

        hideStatus();

        try {
            // Convert Excel to JSON
            const jsonData = await excelToJson(selectedFile, selectedSheet, hasHeader);

            // Format JSON
            const jsonString = prettify
                ? JSON.stringify(jsonData, null, 2)
                : JSON.stringify(jsonData);

            // Copy to clipboard
            await navigator.clipboard.writeText(jsonString);

            // Show success message
            showStatus('Conversion complete! JSON copied to clipboard.', true);

            // Show the JSON with typing effect
            typeJsonOutput(jsonString, prettify);
        } catch (error) {
            console.error('Error converting Excel to JSON:', error);
            showStatus('Error: ' + error.message, false);
        }
    });

    // Copy JSON button click handler
    copyJsonBtn.addEventListener('click', async function() {
        if (currentJsonString) {
            try {
                await navigator.clipboard.writeText(currentJsonString);
                showStatus('JSON copied to clipboard!', true);

                // Change button text temporarily
                const originalText = copyJsonBtn.innerHTML;
                copyJsonBtn.innerHTML = `
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                        <polyline points="20 6 9 17 4 12"></polyline>
                    </svg>
                    Copied!
                `;

                setTimeout(() => {
                    copyJsonBtn.innerHTML = originalText;
                }, 2000);
            } catch (error) {
                console.error('Error copying to clipboard:', error);
                showStatus('Error copying to clipboard.', false);
            }
        }
    });

    // Cancel button click handler
    cancelBtn.addEventListener('click', function() {
        resetForm();
    });

    // Helper function to populate sheet select
    function populateSheetSelect(sheets) {
        sheetSelect.innerHTML = '';

        if (sheets.length === 0) {
            const option = document.createElement('option');
            option.value = '';
            option.textContent = 'No sheets found';
            option.disabled = true;
            sheetSelect.appendChild(option);
            return;
        }

        sheets.forEach(sheet => {
            const option = document.createElement('option');
            option.value = sheet;
            option.textContent = sheet;
            sheetSelect.appendChild(option);
        });

        if (sheets.length > 0) {
            sheetSelect.selectedIndex = 0;
        }
    }

    // Helper function to reset sheet select
    function resetSheetSelect() {
        sheetSelect.innerHTML = '<option value="" disabled selected>First select an Excel file</option>';
    }

    // Helper function to reset the form
    function resetForm() {
        fileInput.value = '';
        fileName.textContent = '';
        fileName.style.display = 'none';
        resetSheetSelect();
        headerRowCheck.checked = true;
        prettifyCheck.checked = true;
        hideStatus();
        resultsContainer.style.display = 'none';
        jsonOutput.textContent = '';
        currentJsonString = '';
    }
});
