function parseExcelDate(serial) {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    return new Date(excelEpoch.getTime() + serial * 24 * 60 * 60 * 1000);
}

function updateBillingMonths(startMonth, endMonth) {
    console.log("startMonth",startMonth,'endMonth',endMonth)
    let startDate = parseExcelDate(startMonth);
    let endDate = parseExcelDate(endMonth);

    if (!startDate && !endDate) return "";

    const formatMonthYear = (date) => {
        return date.toLocaleString('default', { month: 'short', year: 'numeric' }).toUpperCase();
    };

    if (startDate && endDate) {
        const startText = formatMonthYear(startDate);
        const endText = formatMonthYear(endDate);

        if (startText === endText) {
            return `(FOR THE MONTH OF ${startText})`;
        } else {
            return `(FOR THE PERIOD OF ${startText} - ${endText})`;
        }
    } else if (startDate) {
        const startText = formatMonthYear(startDate);
        return `(FOR THE MONTH OF ${startText})`;
    } else {
        return "";
    }
}

function monthYearToExcelSerial(monthYearStr) {
    if (!monthYearStr) return null;

    const [yearStr, monthStr] = monthYearStr.split("-");
    const year = parseInt(yearStr, 10);
    const month = parseInt(monthStr, 10);

    // Correct date: (month - 1)
    const jsDate = new Date(year, month - 1, 1);

    const excelEpoch = new Date(1899, 11, 30);
    const diffInMs = jsDate.getTime() - excelEpoch.getTime();
    let serial = Math.floor(diffInMs / (1000 * 60 * 60 * 24));

    // Adjust for Excel leap year bug
    if (serial >= 61) {
        serial += 1;
    }

    return serial;
}

function convertToExcelSerial(dateInput) {
    if (!dateInput) return null;

    let dateObj;

    if (typeof dateInput === 'string') {
        const [year, month] = dateInput.split('-');
        dateObj = new Date(parseInt(year, 10), parseInt(month, 10) - 1, 1);
    } else if (dateInput instanceof Date) {
        dateObj = dateInput;
    } else {
        return null;
    }

    const excelEpoch = new Date(1899, 11, 30);
    const diffInMs = dateObj.getTime() - excelEpoch.getTime();
    let serialNumber = Math.floor(diffInMs / (1000 * 60 * 60 * 24));

    // Adjust for Excel leap year bug
    if (serialNumber >= 61) {
        serialNumber += 1;
    }

    return serialNumber;
}


function generateInvoice(data) {
    let total = data.pricing.reduce((acc, i) => acc + i, 0);

    document.getElementById("company-name").textContent = data.companyName;
    document.getElementById("company-address").textContent = data.companyAddress;
    document.getElementById("invoice-date").textContent = data.invoiceDate;
    document.getElementById("client-name").textContent = data.client.name;
    document.getElementById("client-location").textContent = data.client.location;
    document.getElementById("bank-name").textContent = data.bankDetails.bankName;
    document.getElementById("bank-ac-no").textContent = `A/C No: ${data.bankDetails.bankAccountNo}`;
    document.getElementById("bank-ifsc-no").textContent = `IFSC: ${data.bankDetails.bankIfscCode}`;
    document.getElementById("bank-pan-no").textContent = `PAN: ${data.bankDetails.bankPanNo}`;
    document.getElementById("signature-name").textContent = data.signature;
    document.getElementById("total-amount").textContent = `₹${total}/-`;
    document.getElementById("total-in-words").textContent = `(Rupees ${convertToWords(total)} Only)`;

    // Update individual pricing fields in the invoice
    data.pricing.forEach((amount, index) => {
        let priceElement = document.getElementById(`pricing-${index + 1}`);
        if (priceElement) priceElement.textContent = `₹${amount}/-`;
    });

    const serialFrom = convertToExcelSerial(data.billingDateFrom);
    const serialTo = convertToExcelSerial(data.billingDateTo);
    
    document.getElementById("bilingMonths").textContent=`${updateBillingMonths(serialFrom,serialTo)}`;

    const outputData = {
        ...data,
        billingDateFrom: monthYearToExcelSerial(data.billingDateFrom),
        billingDateTo: monthYearToExcelSerial(data.billingDateTo),
    };
    window.downloadPDFData = outputData
}


function clearBillingMonths() {
    document.getElementById("billingStart").value = "";
    document.getElementById("billingEnd").value = "";
    document.getElementById("bilingMonths").innerText = "";
}


function convertToWords(num) {
    const ones = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"];
    const teens = ["Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
    const tens = ["", "Ten", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
    const thousands = ["", "Thousand", "Lakh", "Crore"];

    if (num === 0) return "Zero";

    let words = "";

    function convertToWordsHelper(n, place) {
        if (n > 0) {
            if (n < 10) words += ones[n] + " ";
            else if (n >= 11 && n <= 19) words += teens[n - 11] + " ";
            else if (n % 10 === 0) words += tens[n / 10] + " ";
            else {
                words += tens[Math.floor(n / 10)] + " ";
                if (n % 10 > 0) words += ones[n % 10] + " ";
            }
            words += thousands[place] + " ";
        }
    }

    let crore = Math.floor(num / 10000000);
    num %= 10000000;
    let lakh = Math.floor(num / 100000);
    num %= 100000;
    let thousand = Math.floor(num / 1000);
    num %= 1000;
    let hundred = Math.floor(num / 100);
    num %= 100;

    if (crore) convertToWordsHelper(crore, 3);
    if (lakh) convertToWordsHelper(lakh, 2);
    if (thousand) convertToWordsHelper(thousand, 1);
    if (hundred) words += ones[hundred] + " Hundred ";

    if (num > 0) words += "and ";
    convertToWordsHelper(num, 0);

    return words.trim();
}

function downloadPDF() {
    generateHiddenInvoicePDF(window.downloadPDFData)
}


function toggleCompanyFields(enable) {
    const fields = ["companyName", "companyAddress", "bankName", "bankAccountNo", "bankIfscCode", "bankPanNo", "signature"];
    fields.forEach(id => {
        let element = document.getElementById(id);
        if (element) element.disabled = !enable;
    });
}

function validateFormFields() {
    let requiredFields = document.querySelectorAll("#invoiceForm [data-required='true']");
    let firstErrorField = null;

    requiredFields.forEach(field => {
        let parent = field.parentNode;

        // Check if an existing error message is present and remove it
        let existingError = parent.querySelector(".error-text");
        if (existingError) existingError.remove();

        if (!field.value.trim()) {
            field.classList.add("error"); // Highlight field with error

            // Create error message
            let errorMessage = document.createElement("span");
            errorMessage.className = "error-text";
            errorMessage.textContent = "This field is required!";
            errorMessage.style.color = "red";
            errorMessage.style.fontSize = "12px";
            errorMessage.style.display = "block";
            errorMessage.style.marginTop = "5px";

            // Ensure the error message is inserted AFTER the input field
            if (field.nextSibling) {
                field.parentNode.insertBefore(errorMessage, field.nextSibling.nextSibling);
            } else {
                field.parentNode.appendChild(errorMessage);
            }

            if (!firstErrorField) firstErrorField = field; // Set focus on first error
        } else {
            field.classList.remove("error"); // Remove highlight if valid
        }
    });

    if (firstErrorField) {
        firstErrorField.focus();
        return false; // Stop execution if validation fails
    }
    return true; // Allow form submission if all fields are valid
}


// ✅ Function to fetch and process invoice details from the form
function generateManualInvoiceDetailFromForm() {
    try {
        const invoiceData = {
            companyName: document.getElementById("companyName").value.trim() || "",
            companyAddress: document.getElementById("companyAddress").value.trim() || "",
            invoiceDate: document.getElementById("invoiceDate").value.trim() || "",
            client: {
                name: document.getElementById("clientName").value.trim() || "",
                location: document.getElementById("clientLocation").value.trim() || ""
            },
            pricing: [0, 1, 2, 3].map(i => Number(document.getElementById(`pricing${i}`)?.value.trim() || 0)),
            bankDetails: {
                bankName: document.getElementById("bankName").value.trim() || "",
                bankAccountNo: document.getElementById("bankAccountNo").value.trim() || "",
                bankIfscCode: document.getElementById("bankIfscCode").value.trim() || "",
                bankPanNo: document.getElementById("bankPanNo").value.trim() || ""
            },
            signature: document.getElementById("signature").value.trim() || "",
            billingDateFrom: document.getElementById("billingStart").value.trim() || '',
            billingDateTo: document.getElementById("billingEnd").value.trim() || '',
        };

        window.mannualInvoiceData=invoiceData
        generateInvoice(invoiceData); // Process invoice generation

        // Hide the form & show invoice preview
        document.getElementById("invoiceForm").style.display = "none";
        document.getElementById("download-template-wrapper").style.display = "block";

    } catch (error) {
        console.error("Error generating manual invoice:", error);
    }
}

function backToFormDetails() {
    document.getElementById("invoiceForm").style.display = "block";
    document.getElementById("download-template-wrapper").style.display = "none";
}

function backToUploadExcel(){
    document.getElementById("uploadSection").style.display = "block";
    document.getElementById("tableContainer").style.display = "none";
}

// ✅ Attach event listener to the form for validation and submission handling
document.getElementById("invoiceForm").addEventListener("submit", function(event) {
    event.preventDefault(); // Prevent form submission

    if (!validateFormFields()) return; // Validate required fields before proceeding

    generateManualInvoiceDetailFromForm(); // Process invoice if validation passes
});

function openTab(tabId, event) {
    document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
    document.querySelectorAll('.tab').forEach(button => button.classList.remove('active'));

    document.getElementById(tabId).classList.add('active');
    if (event) event.currentTarget.classList.add('active');
}

const fileInput = document.getElementById("fileInput");

fileInput.addEventListener("click", function () {
    this.value = "";
});

function uploadExcel(event) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (e) {
        const workbook = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        const newJsonData = jsonData.filter((item)=>item.length>0)
        if (newJsonData.length > 1) {
            document.getElementById("uploadSection").style.display = "none";
            document.getElementById("tableContainer").style.display = "block";
            window.originalData=newJsonData
            populateInvoiceTable(newJsonData);
        } else {
            alert("The uploaded file is empty.");
        }
    };

    reader.readAsArrayBuffer(file);
}


function excelSerialToDate(serial) {
    let excelEpoch = new Date(1899, 11, 30); // Excel starts from 1899-12-30
    let date = new Date(excelEpoch.getTime() + serial * 86400000); // Convert days to milliseconds
    return date.toLocaleDateString("en-US"); // Format as MM/DD/YYYY
}

function populateInvoiceTable(data) {
    const tableBody = document.getElementById("invoiceTableBody");
    const paginationContainer = document.getElementById("invoicePagination");
    const searchInput = document.getElementById("searchCompany");
    const billingFromInput = document.getElementById("billingFrom");
    const billingToInput = document.getElementById("billingTo");

    let currentPage = 1;
    const rowsPerPage = 5;
    window.filteredData = [...data]; // Store filtered data separately

    function applyFilters() {
        const searchValue = searchInput.value.trim().toLowerCase();
        
        function parseBillingDate(dateStr) {
            if (!dateStr) return null;
            if (typeof dateStr === "number") {
                const excelEpoch = new Date(1899, 11, 30);
                const days = Math.floor(dateStr);
                excelEpoch.setDate(excelEpoch.getDate() + days);
                return excelEpoch;
            }
            if (typeof dateStr === "string") {
                const [month, year] = dateStr.split("/");
                if (!month || !year) return null;
                return new Date(`${year}-${month.padStart(2, '0')}-01`);
            }
            return null;
        }
          
        const fromDate = billingFromInput.value ? parseBillingDate(billingFromInput.value) : null;
        const toDate = billingToInput.value ? parseBillingDate(billingToInput.value) : null;
    
        filteredData = data.filter((row, index) => {
            if (index === 0) return false;
    
            const clientName = (row[8] || "").toLowerCase(); // Client Name
            const rowFromDate = parseBillingDate(row[15]); // Billing Month Start
            const rowToDate = parseBillingDate(row[16]); // Billing Month End
    
            const matchesSearch = !searchValue || clientName.includes(searchValue);
            const matchesFromDate = !fromDate || !rowFromDate || rowFromDate >= fromDate;
            const matchesToDate = !toDate || !rowToDate || rowToDate <= toDate;
    
            return matchesSearch && matchesFromDate && matchesToDate;
        });
        
        currentPage = 1;
        renderTable(currentPage);
        renderPagination();
    }
    

    function renderTable(page) {
        tableBody.innerHTML = "";
        const totalRows = filteredData.length;
        if (totalRows === 0) {
            tableBody.innerHTML = "<tr><td colspan='18'>No data found</td></tr>";
            paginationContainer.innerHTML = ""; 
            return;
        }
    
        let start = (page - 1) * rowsPerPage;
        let end = Math.min(start + rowsPerPage, totalRows);
    
        for (let i = start; i < end; i++) {
            let row = document.createElement("tr");
            row.innerHTML = `
                <td class="custom-fixed-column">${filteredData[i][0] || "-"}</td> <!-- Company Name -->
                <td class="custom-th">${filteredData[i][1] || "-"}</td> <!-- Company Address -->
                <td class="custom-th">${filteredData[i][2] || "-"}</td> <!-- Bank Name -->
                <td class="custom-th">${filteredData[i][3] || "-"}</td> <!-- Bank Account No -->
                <td class="custom-th">${filteredData[i][4] || "-"}</td> <!-- Bank IFSC Code -->
                <td class="custom-th">${filteredData[i][5] || "-"}</td> <!-- Bank PAN No -->
                <td class="custom-th">${filteredData[i][6] || "-"}</td> <!-- Signature -->
                <td class="custom-th">${excelSerialToDate(filteredData[i][7]) || "-"}</td> <!-- Invoice Date -->
                <td class="custom-th">${filteredData[i][8] || "-"}</td> <!-- Client Name -->
                <td class="custom-th">${filteredData[i][9] || "-"}</td> <!-- Client Location -->
                <td class="custom-th">${filteredData[i][10] || "-"}</td> <!-- EPFO charges -->
                <td class="custom-th">${filteredData[i][11] || "-"}</td> <!-- ESIC charges -->
                <td class="custom-th">${filteredData[i][12] || "-"}</td> <!-- Postal charges -->
                <td class="custom-th">${filteredData[i][13] || "-"}</td> <!-- Other charges -->
                <td class="custom-th">${filteredData[i][14] || "-"}</td> <!-- Total Amount -->
                <td class="custom-th">${excelSerialToDate(filteredData[i][15]) || "-"}</td> <!-- Billing Month start -->
                <td class="custom-th">${excelSerialToDate(filteredData[i][16]) || "-"}</td> <!-- Billing Month end -->
                <td class="custom-action-column">
                    <span class="download-btn"><i class="fa-solid fa-download"></i></span>
                </td>
            `;

            row.querySelector(".download-btn").addEventListener("click", () => {
                generateHiddenInvoicePDF({
                    companyName: filteredData[i][0],
                    companyAddress: filteredData[i][1],
                    bankDetails: {
                        bankName: filteredData[i][2],
                        bankAccountNo: filteredData[i][3],
                        bankIfscCode: filteredData[i][4],
                        bankPanNo: filteredData[i][5]
                    },
                    signature: filteredData[i][6],
                    invoiceDate: excelSerialToDate(filteredData[i][7]),
                    client: {
                        name: filteredData[i][8],
                        location: filteredData[i][9]
                    },
                    pricing: [
                        parseFloat(filteredData[i][10]) || 0,
                        parseFloat(filteredData[i][11]) || 0,
                        parseFloat(filteredData[i][12]) || 0,
                        parseFloat(filteredData[i][13]) || 0,
                    ],
                    totalCosting: filteredData[i][14],
                    billingDateFrom: filteredData[i][15],
                    billingDateTo: filteredData[i][16],
                });
            });
    
            tableBody.appendChild(row);
        }
    }
    
    

    function renderPagination() {
        paginationContainer.innerHTML = "";
        const totalPages = Math.ceil(filteredData.length / rowsPerPage);
        if (totalPages <= 1) return;

        let prevButton = document.createElement("button");
        prevButton.innerText = "«";
        prevButton.classList.add("page-btn");
        prevButton.disabled = currentPage === 1;
        prevButton.addEventListener("click", () => {
            if (currentPage > 1) {
                currentPage--;
                renderTable(currentPage);
                renderPagination();
            }
        });
        paginationContainer.appendChild(prevButton);
    
        // Logic to display page numbers dynamically
        if (totalPages <= 5) {
            // If pages are 5 or less, show all pages
            for (let i = 1; i <= totalPages; i++) {
                addPageButton(i);
            }
        } else {
            addPageButton(1); // First page
            
            if (currentPage > 3) addEllipsis();
    
            for (let i = Math.max(2, currentPage - 1); i <= Math.min(totalPages - 1, currentPage + 1); i++) {
                addPageButton(i);
            }
    
            if (currentPage < totalPages - 2) addEllipsis();
    
            addPageButton(totalPages); // Last page
        }
    
        // Create Next Button
        let nextButton = document.createElement("button");
        nextButton.innerText = "»";
        nextButton.classList.add("page-btn");
        nextButton.disabled = currentPage === totalPages;
        nextButton.addEventListener("click", () => {
            if (currentPage < totalPages) {
                currentPage++;
                renderTable(currentPage);
                renderPagination();
            }
        });
        paginationContainer.appendChild(nextButton);
    }

    function addPageButton(page) {
        let pageButton = document.createElement("button");
        pageButton.innerText = page;
        pageButton.classList.add("page-btn");
        if (page === currentPage) pageButton.classList.add("active");
        pageButton.addEventListener("click", () => {
            currentPage = page;
            renderTable(currentPage);
            renderPagination();
        });
        paginationContainer.appendChild(pageButton);
    }

    function addEllipsis() {
        let ellipsis = document.createElement("span");
        ellipsis.innerText = "...";
        ellipsis.classList.add("ellipsis");
        paginationContainer.appendChild(ellipsis);
    }


    searchInput.addEventListener("input", applyFilters);
    billingFromInput.addEventListener("change", applyFilters);
    billingToInput.addEventListener("change", applyFilters);

    applyFilters();
}


function generateHiddenInvoicePDF(data, isBatchMode = false) {
    return new Promise((resolve) => {
        const { jsPDF } = window.jspdf;

        let hiddenDiv = document.createElement("div");
        hiddenDiv.style.position = "absolute";
        hiddenDiv.style.left = "-9999px";
        document.body.appendChild(hiddenDiv);

        let invoiceElement = document.getElementById("download-template");
        if (!invoiceElement) {
            resolve(null);
            return;
        }

        hiddenDiv.innerHTML = invoiceElement.outerHTML;

        setTimeout(() => {
            let invoice = hiddenDiv.querySelector("#download-template");
            if (!invoice) {
                document.body.removeChild(hiddenDiv);
                resolve(null);
                return;
            }

            function safeUpdate(selector, value) {
                let element = invoice.querySelector(selector);
                if (element) {
                    element.textContent = value || "";
                } 
            }
            // ** Update invoice details **
            safeUpdate("#company-name", data.companyName);
            safeUpdate("#company-address", data.companyAddress);
            safeUpdate("#invoice-date", data.invoiceDate);
            safeUpdate("#client-name", data.client.name);
            safeUpdate("#client-location", data.client.location);
            safeUpdate("#bank-name", data.bankDetails.bankName);
            safeUpdate("#bank-ac-no", `A/C No: ${data.bankDetails.bankAccountNo}`);
            safeUpdate("#bank-ifsc-no", `IFSC: ${data.bankDetails.bankIfscCode}`);
            safeUpdate("#bank-pan-no", `PAN: ${data.bankDetails.bankPanNo}`);
            safeUpdate("#signature-name", data.signature);
            safeUpdate("#bilingMonths", `${updateBillingMonths(data.billingDateFrom,data.billingDateTo)}`);
            let total = 0;
            data.pricing.forEach((amount, index) => {
                safeUpdate(`#pricing-${index + 1}`, `₹${amount}/-`);
                total += parseFloat(amount) || 0;
            });

            safeUpdate("#total-amount", `₹${total}/-`);
            safeUpdate("#total-in-words", `(Rupees ${convertToWords(total)} Only)`);

            // ** Ensure A4 fit **
            invoice.style.transform = "scale(0.85)";
            invoice.style.transformOrigin = "top left";
            invoice.style.width = "210mm";
            invoice.style.height = "auto";

            let table = invoice.querySelector(".invoice-table");
            if (table) {
                table.style.fontSize = "13px";
                table.style.width = "100%";
            }

            html2canvas(invoice, { 
                scale: 1.5,  // Decrease scale to reduce resolution
                backgroundColor: null, 
            }).then((canvas) => {
                const imgData = canvas.toDataURL("image/png");
                const pdf = new jsPDF("p", "mm", "a4");
                const pageWidth = 210;
                const pageHeight = 297;
            
                let imgWidth = pageWidth;
                let imgHeight = (canvas.height * imgWidth) / canvas.width;
            
                if (imgHeight > pageHeight) {
                    imgHeight = pageHeight;
                }
            
                pdf.addImage(imgData, "PNG", 0, 0, imgWidth, imgHeight);
                if (isBatchMode) {
                    resolve(pdf.output("blob"));
                } else {
                    pdf.save(`invoice_${data.client.name}-${data.invoiceDate}.pdf`);
                    resolve(null);
                }
            
                document.body.removeChild(hiddenDiv);
            });
            
        }, 300);
    });
}

async function downloadAllInvoicesAsZip() {
    await loadJSZip();

    let zip = new JSZip();
    const batchSize = 10; // Process 10 at a time
    const flushThreshold = 50; // Save ZIP every 50 invoices

    if (!window.filteredData || window.filteredData.length < 1) {  
        alert("No invoices found!");
        return;
    }

    document.body.style.pointerEvents = "none";
    document.body.style.overflow = "hidden";

    const progressOverlay = document.createElement("div");
    progressOverlay.id = "progress-overlay";
    progressOverlay.innerHTML = '<div id="progress-container"><div id="progress-bar"></div><span id="progress-text">Starting...</span></div>';
    document.body.appendChild(progressOverlay);

    let totalRows = window.filteredData.length;
    let completedRows = 0;
    let zipCount = 0; // Track separate ZIP files

    function updateProgress() {
        let percent = Math.round((completedRows / totalRows) * 100);
        document.getElementById("progress-bar").style.width = percent + "%";
        document.getElementById("progress-text").textContent = `Processing ${completedRows} of ${totalRows} invoices...`;
    }

    async function processBatch(startIndex, endIndex) {
        const batchPromises = [];
    
        for (let i = startIndex; i < endIndex; i++) {
            batchPromises.push((async () => {
                try {
                    const row = window.filteredData[i];
                    if (!row || row.length < 14) return;
                    const dataToSet = {
                        companyName: row[0],
                        companyAddress: row[1],
                        invoiceDate: excelSerialToDate(row[7]) ,
                        client: { name: row[8], location: row[9] },
                        pricing: [parseFloat(row[10]) || 0, parseFloat(row[11]) || 0, parseFloat(row[12]) || 0, parseFloat(row[13]) || 0],
                        bankDetails: { bankName: row[2], bankAccountNo: row[3], bankIfscCode: row[4], bankPanNo: row[5] },
                        signature: row[6],
                        billingDateFrom:row[15],
                        billingDateTo:row[16],
                    };

                    const pdfBlob = await generateHiddenInvoicePDF(dataToSet, true);
    
                    if (!pdfBlob || !(pdfBlob instanceof Blob)) {
                        return;
                    }
    
                    let uniqueId = new Date().getTime() + Math.floor(Math.random() * 1000);
                    let fileName = `invoice_${dataToSet.client.name}_${dataToSet.invoiceDate}_${uniqueId}.pdf`;
                    fileName = fileName.replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_.-]/g, '');
                    zip.file(fileName, pdfBlob);
                    completedRows++;
                    updateProgress();
    
                } catch (error) {
                    console.error("Error generating PDF for row:", i, error);
                }
            })());
        }
    
        await Promise.allSettled(batchPromises);
        await new Promise(resolve => setTimeout(resolve, 100));
    }

    async function processAllBatches() {
        for (let i = 0; i < totalRows; i += batchSize) {
            await processBatch(i, Math.min(i + batchSize, totalRows));
    
            if ((i + batchSize) % flushThreshold === 0 || i + batchSize >= totalRows) {  
                const partialZip = await zip.generateAsync({ type: "blob" });

                saveAs(partialZip, `Invoices_Part_${++zipCount}.zip`);

                zip = new JSZip(); // Reset ZIP for next batch
            }

            await new Promise(resolve => setTimeout(resolve, 100));
        }

        document.getElementById("progress-text").textContent = "Download Complete!";
        setTimeout(() => {
            document.body.removeChild(progressOverlay);
            document.body.style.pointerEvents = "auto";
            document.body.style.overflow = "auto";
        }, 2000);
    }

     setTimeout(processAllBatches, 0);
}

// ✅ Ensure JSZip is Loaded Before Use
async function loadJSZip() {
    if (typeof JSZip === "undefined") {
        await new Promise((resolve, reject) => {
            const script = document.createElement("script");
            script.src = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";
            script.onload = resolve;
            script.onerror = () => reject(new Error("Failed to load JSZip!"));
            document.head.appendChild(script);
        });
    } 

}

document.getElementById('downloadExcelBtn').addEventListener('click', function() {
    const link = document.createElement('a');
    link.href = 'files/CopyBillGeneratorSheet.xlsx'; // path to your file
    link.download = 'dummyBillGenerator.xlsx';   // suggested file name
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
});