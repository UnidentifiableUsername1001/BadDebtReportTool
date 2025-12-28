        const sheet1Input = document.getElementById('sheet1');
        const sheet2Input = document.getElementById('sheet2');
        const sheet3Input = document.getElementById('sheet3');
        const generateBtn = document.getElementById('generate-report-btn');
        const errorMessage = document.getElementById('error-message');
        const resultsSection = document.getElementById('results-section');
        const loader = document.getElementById('loader');

        // Helper function to format numbers as currency
        const formatCurrency = (value) => {
            return new Intl.NumberFormat('en-GB', { style: 'currency', currency: 'GBP' }).format(value);
        };

        // Function to parse a CSV file and return a promise
        const parseFile = (file, options = {}) => {
            return new Promise((resolve, reject) => {
                if (!file) {
                    reject(new Error("File not selected."));
                    return;
                }
                const config = {
                    header: true,
                    skipEmptyLines: true,
                    transformHeader: header => header.trim(),
                    ...options, // Allow overriding default options
                };
                Papa.parse(file, {
                    ...config,
                    complete: (results) => resolve(results.data),
                    error: (error) => reject(error),
                });
            });
        };

        // Main function to generate the report
        const generateReport = async () => {
            // 1. Clear previous state and show loader
            errorMessage.textContent = '';
            resultsSection.classList.add('hidden');
            loader.classList.remove('hidden');
            generateBtn.disabled = true;

            // 2. Check if all files are selected
            if (!sheet1Input.files[0] || !sheet2Input.files[0] || !sheet3Input.files[0]) {
                errorMessage.textContent = 'Please upload all three CSV files.';
                loader.classList.add('hidden');
                generateBtn.disabled = false;
                return;
            }

            try {
                // 3. Parse all files concurrently
                const [sheet1Data, sheet2Data, sheet3Data] = await Promise.all([
                    parseFile(sheet1Input.files[0]),
                    parseFile(sheet2Input.files[0], { header: false }), // Parse Sheet 2 by position
                    parseFile(sheet3Input.files[0]),
                ]);
                
                // 4. Create lookup maps for efficient data retrieval
                
                // Build Sheet 2 map using column index (A=0, BF=57)
                const sheet2Map = new Map();
                // Start from 1 to skip the header row in Sheet 2
                for (let i = 1; i < sheet2Data.length; i++) {
                    const row = sheet2Data[i];
                    // Safety check: Ensure row has enough columns
                    if (row && row.length > 0) {
                        const id = row[0] ? String(row[0]).trim() : null; // Column A
                        const status = row[57] ? String(row[57]).trim() : 'N/A'; // Column BF
                        if (id) {
                            sheet2Map.set(id, status);
                        }
                    }
                }

                const sheet3Map = new Map(sheet3Data.map(row => [row['Auction Id'], {
                    recoveredPrincipal: parseFloat((row['Recovered Principal'] || '0').replace(/,/g, '')),
                    recoveredInterest: parseFloat((row['Recovered Interest'] || '0').replace(/,/g, '')),
                }]));

                // 5. Process the data
                const badDebtLoans = sheet1Data.filter(row => row['Loan status'] === 'Bad Debt');

                let totalPrincipalRemaining = 0;
                let totalPrincipalRecovered = 0;
                let totalInterestRecovered = 0;
                let totalOutstanding = 0;

                const reportData = badDebtLoans.map(loan => {
                    const loanId = loan['Loan ID'] ? String(loan['Loan ID']).trim() : null;

                    const recoveryStatus = sheet2Map.get(loanId) || 'N/A';
                    const recoveryData = sheet3Map.get(loanId) || { recoveredPrincipal: 0, recoveredInterest: 0 };
                    const principalRemaining = parseFloat((loan['Principal remaining'] || '0').replace(/,/g, ''));
                    const principalRecovered = recoveryData.recoveredPrincipal;
                    const interestRecovered = recoveryData.recoveredInterest;
                    const outstandingBalance = principalRemaining - principalRecovered - interestRecovered;
                    
                    totalPrincipalRemaining += principalRemaining;
                    totalPrincipalRecovered += principalRecovered;
                    totalInterestRecovered += interestRecovered;
                    totalOutstanding += outstandingBalance;

                    return {
                        'Loan ID': loanId,
                        'Principal at Default': principalRemaining,
                        'Recovery Status': recoveryStatus,
                        'Principal Recovered': principalRecovered,
                        'Interest Recovered': interestRecovered,
                        'Outstanding Balance': outstandingBalance,
                    };
                });

                // Sort: Defaulted cases first
                // Logic uses "defaulted" as the keyword (internal value)
                reportData.sort((a, b) => {
                    if (a['Recovery Status'] === 'defaulted' && b['Recovery Status'] !== 'defaulted') return -1;
                    if (b['Recovery Status'] === 'defaulted' && a['Recovery Status'] !== 'defaulted') return 1;
                    return 0;
                });

                // 6. Calculate summary metrics for 'defaulted' loans
                const defaultedLoans = reportData.filter(loan => loan['Recovery Status'] === 'defaulted');
                const defaultedCasesCount = defaultedLoans.length;
                const totalOutstandingForDefaulted = defaultedLoans.reduce((sum, loan) => sum + loan['Outstanding Balance'], 0);
                
                // 7. Display summary and table
                displaySummary(defaultedCasesCount, totalOutstandingForDefaulted);
                displayTable(reportData, {totalPrincipalRemaining, totalPrincipalRecovered, totalInterestRecovered, totalOutstanding});
                resultsSection.classList.remove('hidden');

            } catch (error) {
                console.error("Error generating report:", error);
                errorMessage.textContent = 'An error occurred while processing the files. Please check the console (F12) for details.';
            } finally {
                // 8. Hide loader and re-enable button
                loader.classList.add('hidden');
                generateBtn.disabled = false;
            }
        };

        const displaySummary = (count, total) => {
            document.getElementById('defaulted-cases-count').textContent = count;
            document.getElementById('total-outstanding-balance').textContent = formatCurrency(total);
        };
        
        const displayTable = (data, totals) => {
            const table = document.getElementById('report-table');
            const thead = table.querySelector('thead');
            const tbody = table.querySelector('tbody');
            const tfoot = table.querySelector('tfoot');
            
            thead.innerHTML = '';
            tbody.innerHTML = '';
            tfoot.innerHTML = '';

            if (data.length === 0) {
                tbody.innerHTML = '<tr><td colspan="6" style="text-align:center; padding: 20px;">No bad debt loans found in the provided file.</td></tr>';
                return;
            }

            const headers = Object.keys(data[0]);
            const headerRow = document.createElement('tr');
            headers.forEach(headerText => {
                const th = document.createElement('th');
                th.textContent = headerText;
                headerRow.appendChild(th);
            });
            thead.appendChild(headerRow);
            
            data.forEach(rowData => {
                const row = document.createElement('tr');
                
                // Highlight row if status is 'defaulted' (which displays as 'Open Case')
                if (rowData['Recovery Status'] === 'defaulted') {
                    row.classList.add('row-open-case');
                }

                headers.forEach(header => {
                    const cell = document.createElement('td');
                    let value = rowData[header];

                    // --- FORMATTING LOGIC FOR DISPLAY ONLY ---
                    if (header === 'Recovery Status') {
                        if (value === 'defaulted') {
                            value = 'Open Case';
                        } else if (value === 'closed_case') {
                            value = 'Closed Case';
                        }
                    }

                    if (['Principal at Default', 'Principal Recovered', 'Interest Recovered', 'Outstanding Balance'].includes(header)) {
                         cell.textContent = formatCurrency(value);
                         cell.classList.add('text-right');
                    } else {
                         cell.textContent = value;
                    }
                    row.appendChild(cell);
                });
                tbody.appendChild(row);
            });

            const footerRow = document.createElement('tr');
            footerRow.innerHTML = `
                <td><strong>Totals</strong></td>
                <td class="text-right">${formatCurrency(totals.totalPrincipalRemaining)}</td>
                <td></td>
                <td class="text-right">${formatCurrency(totals.totalPrincipalRecovered)}</td>
                <td class="text-right">${formatCurrency(totals.totalInterestRecovered)}</td>
                <td class="text-right">${formatCurrency(totals.totalOutstanding)}</td>
            `;
            tfoot.appendChild(footerRow);
        };

        generateBtn.addEventListener('click', generateReport);