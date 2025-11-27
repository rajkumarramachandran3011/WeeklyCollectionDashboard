document.addEventListener('DOMContentLoaded', () => {
    // --- START: Firebase Configuration ---
    // The firebaseConfig object is now loaded from the untracked config.js file
    // Initialize Firebase
    firebase.initializeApp(firebaseConfig);
    const db = firebase.firestore(); // Get a reference to the Firestore database
    // --- END: Firebase Configuration ---

    const addCustomerBtn = document.getElementById('add-customer-btn');
    const modal = document.getElementById('add-customer-modal');
    const closeBtn = document.querySelector('.close-btn');
    const addCustomerForm = document.getElementById('add-customer-form');
    const customerGridBody = document.querySelector('#customer-grid tbody');
    const dayFilters = document.querySelector('.day-filters');
    const searchInput = document.getElementById('search-input');
    const prevWeekBtn = document.getElementById('prev-week-btn');
    const nextWeekBtn = document.getElementById('next-week-btn');
    const weekRangeEl = document.getElementById('week-range');
    const selectedDateDisplay = document.getElementById('selected-date-display');

    // New DOM elements for Transaction History Modal
    const transactionHistoryModal = document.getElementById('transaction-history-modal');
    const transactionHistoryCloseBtn = document.querySelector('.transaction-history-close-btn');
    const historyCustomerName = document.getElementById('history-customer-name');
    const historyTotalPayable = document.getElementById('history-total-payable');
    const historyBalanceAmount = document.getElementById('history-balance-amount');
    const historyGridBody = transactionHistoryModal.querySelector('#transaction-history-table-body');

    let currentDate = new Date();
    let customers = [];

    // Function to update the displayed week range
    const updateWeekRange = () => {
        const startOfWeek = new Date(currentDate);
        startOfWeek.setDate(currentDate.getDate() - currentDate.getDay()); // Go to Sunday
        const endOfWeek = new Date(startOfWeek);
        endOfWeek.setDate(startOfWeek.getDate() + 6); // Go to Saturday

        const options = { year: 'numeric', month: 'short', day: 'numeric' };
        weekRangeEl.textContent = `${startOfWeek.toLocaleDateString('en-US', options)} - ${endOfWeek.toLocaleDateString('en-US', options)}`;
    };

    // Function to update the selected date display
    const updateSelectedDate = (dayName) => {
        const today = new Date(currentDate);
        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        const currentDayIndex = today.getDay();
        const targetDayIndex = days.indexOf(dayName);

        if (targetDayIndex !== -1) {
            const diff = targetDayIndex - currentDayIndex;
            today.setDate(today.getDate() + diff);
        }
        selectedDateDisplay.textContent = `Selected Day: ${today.toLocaleDateString('en-US', { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' })}`;
    };

    // Load customers from Firestore
    const loadCustomers = () => {
        console.log("Setting up Firestore listener...");
        db.collection("customers").onSnapshot((snapshot) => {
            console.log("Received snapshot from Firestore. Number of documents:", snapshot.size);
            customers = snapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
            console.log("Loaded customers:", customers);
            customers.sort((a, b) => a.id.localeCompare(b.id)); // Keep customers sorted
            filterAndRender(); // Re-render the grid whenever data changes
        }, (error) => {
            console.error("Firestore snapshot error: ", error);
            alert("Error connecting to the database. Please check the console for details.");
        });
    };

    // Preselect current day filter
    const preselectCurrentDayFilter = () => {
        const today = new Date();
        // If we are just loading the page, we might not have customers yet.
        // The `loadCustomers` function will call `filterAndRender` once data arrives.
        // So, we only need to set the active button and date display here.
        if (customers.length === 0) return;

        const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
        const currentDayName = days[today.getDay()];

        // Remove active class from any initially active button (like Sunday from index.html)
        const initiallyActiveButton = document.querySelector('.day-filter.active');
        if (initiallyActiveButton) {
            initiallyActiveButton.classList.remove('active');
        }

        // Find and activate the button for the current day
        const currentDayButton = document.querySelector(`.day-filter[data-day="${currentDayName}"]`);
        if (currentDayButton) {
            currentDayButton.classList.add('active');
        } else {
            // Fallback to Sunday if current day button not found (shouldn't happen with full day list)
            document.querySelector('.day-filter[data-day="Sunday"]').classList.add('active');
        }
        updateSelectedDate(currentDayName); // Update the selected date display for the preselected day
    };

    const getNextCustomerId = () => {
        const lastCustomer = customers.length > 0 ? customers[customers.length - 1] : null;
        if (!lastCustomer || !lastCustomer.id) {
            return 'CUST001';
        }
        const lastId = parseInt(lastCustomer.id.replace('CUST', ''));
        const nextId = lastId + 1;
        return 'CUST' + nextId.toString().padStart(3, '0');
    };

    const getWeekId = (date) => {
        const firstDay = new Date(date); // Create a new Date object to avoid modifying the original
        firstDay.setDate(date.getDate() - date.getDay());
        return firstDay.toISOString().split('T')[0];
    }

    const renderGrid = (customerData) => {
        const weekId = getWeekId(new Date(currentDate));
        customerGridBody.innerHTML = '';
        for (const customer of customerData) {
            try {
                const payment = customer.paymentHistory ? customer.paymentHistory[weekId] : null;
                const paymentStatus = payment ? payment.status : 'Pending';
                const lastPaidAmount = payment ? payment.amount : 0;
                const paymentMode = payment ? payment.mode : 'Cash';

                // Provide defaults for potentially missing data
                const balanceAmount = customer.balanceAmount ?? 0;
                const totalPayableAmount = customer.totalPayableAmount ?? 0;
                const customerName = customer.name || 'N/A';
                const customerId = customer.id || 'N/A';
                const customerDay = customer.day || 'N/A';
                const customerAddress = customer.address || '';

                const row = document.createElement('tr');
                row.dataset.customerId = customerId; // Add data-customer-id to the row

                row.innerHTML = `
                    <td data-label="Name" class="customer-name-cell">
                        <div class="customer-name">${customerName}</div>
                        <div class="customer-id">${customerId}</div>
                    </td>
                    <td data-label="Day" style="display: none;">
                        <div>${customerDay}</div>
                        <div class="customer-address">${customerAddress}</div>
                    </td>
                    <td data-label="Balance Amount" class="balance-amount-cell">
                        <div class="balance-amount">₹${balanceAmount.toLocaleString('en-IN')}</div>
                        <div class="total-payable">of ₹${totalPayableAmount.toLocaleString('en-IN')}</div>
                    </td>
                    <td data-label="Amount Paid">
                        <div class="amount-paid-container">
                            <input type="number" class="amount-paid-input" value="${paymentStatus === 'Paid' ? lastPaidAmount : ''}" ${paymentStatus === 'Paid' ? 'disabled' : ''} min="1" max="99999">
                            <select class="payment-mode-select" ${paymentStatus === 'Paid' ? 'disabled' : ''}>
                                <option value="Cash" ${paymentMode === 'Cash' ? 'selected' : ''}>Cash</option>
                                <option value="UPI" ${paymentMode === 'UPI' ? 'selected' : ''}>UPI</option>
                            </select>
                        </div>
                    </td>
                    <td data-label="Actions">
                        ${paymentStatus === 'Paid' ? '<button class="edit-pay-btn"><i class="fas fa-edit"></i> Edit</button>' : '<button class="pay-btn pay-btn-large"><i class="fas fa-money-bill-wave"></i> Pay</button>'}
                    </td>
                    <td data-label="Payment Status">
                        ${paymentStatus === 'Paid' ? `<span class="paid-status">Paid</span><div class="paid-date">${new Date(payment.paymentDate).toLocaleString()}</div>` : '<span class="not-paid-status">Not Paid</span>'}
                    </td>
                `;
                customerGridBody.appendChild(row);
            } catch (error) {
                console.error("Could not render customer row. Data might be corrupt:", customer, error);
            }
        }
    };

    // Function to show/edit customer details in the add-customer-modal
    const showCustomerDetailsModal = (customer) => {
        document.querySelector('#add-customer-modal h2').textContent = 'Edit Customer';
        document.querySelector('#add-customer-form button[type="submit"]').textContent = 'Update Customer';

        document.getElementById('customer-id').value = customer.id;
        document.getElementById('name').value = customer.name;
        document.getElementById('phone').value = customer.phone;
        document.getElementById('address').value = customer.address;
        document.getElementById('day').value = customer.day;
        document.getElementById('total-payable-amount').value = customer.totalPayableAmount;
        // Assuming numberOfInstallments is part of customer object
        document.getElementById('number-of-installments').value = customer.numberOfInstallments || '';
        document.getElementById('account-opening-date').value = customer.accountOpeningDate;

        editingCustomerIndex = customers.findIndex(c => c.id === customer.id);
        deleteCustomerBtn.style.display = 'block'; // Show delete button when editing
        modal.style.display = 'block';
    };

    // Function to show transaction history in the transaction-history-modal
    const showTransactionHistoryModal = (customer) => {
        historyCustomerName.textContent = customer.name;
        historyTotalPayable.textContent = `₹${customer.totalPayableAmount.toLocaleString('en-IN')}`;
        historyBalanceAmount.textContent = `₹${customer.balanceAmount.toLocaleString('en-IN')}`;
        historyGridBody.innerHTML = ''; // Clear previous history

        let totalPaidAmount = 0; // Initialize total paid amount

        if (customer.paymentHistory && Object.keys(customer.paymentHistory).length > 0) {
            const sortedWeeks = Object.keys(customer.paymentHistory).sort((a, b) => new Date(a) - new Date(b));
            let weekNumber = 1; // Initialize week number

            sortedWeeks.forEach(weekId => {
                const payment = customer.paymentHistory[weekId];
                const row = document.createElement('tr');
                // Format weekId to a more readable date
                const weekDate = new Date(weekId).toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });

                row.innerHTML = `
                    <td data-label="Week No.">${weekNumber++}</td> <!-- Week Number -->
                    <td data-label="Week">${weekDate}</td> <!-- Week Date -->
                    <td data-label="Amount Paid">₹${payment.amount.toLocaleString('en-IN')} (${payment.mode})</td> <!-- Amount Paid -->
                `;
                historyGridBody.appendChild(row);

                totalPaidAmount += payment.amount; // Add to total paid amount
            });
        } else {
            const row = document.createElement('tr');
            // colspan should be 3 now as there are 3 columns
            row.innerHTML = `<td colspan="3" style="text-align: center;">No transaction history available.</td>`;
            historyGridBody.appendChild(row);
        }

        // Update the total paid amount display
        document.getElementById('total-paid-amount').textContent = `₹${totalPaidAmount.toLocaleString('en-IN')}`;

        transactionHistoryModal.style.display = 'block';
    };

    let editingCustomerIndex = null;

    // Event Listeners for week navigation
    prevWeekBtn.addEventListener('click', () => {
        currentDate.setDate(currentDate.getDate() - 7);
        updateWeekRange();
        filterAndRender();
    });

    nextWeekBtn.addEventListener('click', () => {
        currentDate.setDate(currentDate.getDate() + 7);
        updateWeekRange();
        filterAndRender();
    });

            

            

                const filterAndRender = () => {
                    const activeDayFilterButton = document.querySelector('.day-filter.active');
                    const dayFilter = activeDayFilterButton ? activeDayFilterButton.dataset.day : 'All';
                    const searchTerm = searchInput.value.toLowerCase();

                    let filteredCustomers = customers;

                    if (dayFilter !== 'All') {
                        filteredCustomers = filteredCustomers.filter(c => c.day === dayFilter);
                    }

                    if (searchTerm) {
                        filteredCustomers = filteredCustomers.filter(c =>
                            c.name.toLowerCase().includes(searchTerm) ||
                            c.phone.toLowerCase().includes(searchTerm) ||
                            (c.address && c.address.toLowerCase().includes(searchTerm))
                        );
                    }

                    // Filter based on account opening date
                    filteredCustomers = filteredCustomers.filter(c => {
                        if (!c.accountOpeningDate) {
                            return true; // If no account opening date is set, always show the customer.
                        }
                        const accountOpeningDate = new Date(c.accountOpeningDate);
                        // Set hours to 0 to compare dates only
                        accountOpeningDate.setHours(0, 0, 0, 0);
                        const currentDateOnly = new Date(currentDate);
                        currentDateOnly.setHours(0, 0, 0, 0);

                        return currentDateOnly >= accountOpeningDate;
                    });

                    renderGrid(filteredCustomers);
                };

            

                                const deleteCustomerBtn = document.getElementById('delete-customer-btn');

            

                

            

                                addCustomerBtn.addEventListener('click', () => {

            

                    editingCustomerIndex = null;

            

                    document.querySelector('#add-customer-modal h2').textContent = 'Add New Customer';

            

                    document.querySelector('#add-customer-form button[type="submit"]').textContent = 'Add Customer'; // Ensure button text is "Add Customer"

            

                    addCustomerForm.reset();

            

                    document.getElementById('customer-id').value = getNextCustomerId();

            

                    const selectedDay = document.querySelector('.day-filter.active').dataset.day;

            

                    document.getElementById('day').value = selectedDay;

            

                    deleteCustomerBtn.style.display = 'none'; // Hide delete button when adding new customer

            

                    modal.style.display = 'block';

            

                });

            

                closeBtn.addEventListener('click', () => {

                    modal.style.display = 'none';

                });

            

                deleteCustomerBtn.addEventListener('click', () => {

                    if (editingCustomerIndex !== null) {

                        const customerToDelete = customers[editingCustomerIndex];
                        if (confirm('Are you sure you want to delete this customer? This action cannot be undone.')) {
                            // Delete from Firestore
                            db.collection("customers").doc(customerToDelete.id).delete().then(() => {
                                console.log("Customer deleted successfully");
                            }).catch(error => {
                                console.error("Error removing document: ", error);
                                alert("Error deleting customer. Check console for details.");
                            });

                            editingCustomerIndex = null;
                            modal.style.display = 'none';

                        }

                    }

                });

            

                window.addEventListener('click', (e) => {

                    if (e.target === modal) {

                        modal.style.display = 'none';

                    }

                });

            

                addCustomerForm.addEventListener('submit', (e) => {

                    e.preventDefault();

                    const name = document.getElementById('name').value;

                    const phone = document.getElementById('phone').value;

            

                    if (editingCustomerIndex === null) { // Only check for duplicates when adding a new customer

                        const isDuplicate = customers.some(customer => customer.name === name && customer.phone === phone);

                        if (isDuplicate) {

                            alert('A customer with the same name and phone number already exists.');

                            return;

                        }

                    }

            

                    const newCustomer = {

                        // id is now managed by Firestore, but we can keep it for display if we want.

                        name: name,

                        phone: phone,

                        address: document.getElementById('address').value,

                        day: document.getElementById('day').value,

            

                        totalPayableAmount: parseInt(document.getElementById('total-payable-amount').value),

                        balanceAmount: parseInt(document.getElementById('total-payable-amount').value), // Initially set balance to total payable

                        amountPaid: 0, // Reset amount paid for new customer

                        lastPaidAmount: 0,

                        accountOpeningDate: document.getElementById('account-opening-date').value,

                        paymentHistory: {}

                    };

            

                    if (editingCustomerIndex !== null) {

                        const customerToUpdate = customers[editingCustomerIndex];
                        // Update in Firestore
                        db.collection("customers").doc(customerToUpdate.id).set(newCustomer, { merge: true }).then(() => {
                            console.log("Customer updated successfully");
                        }).catch(error => console.error("Error updating document: ", error));
                        editingCustomerIndex = null;

                    } else {
                        const customerId = document.getElementById('customer-id').value;
                        // Add to Firestore with a specific ID
                        db.collection("customers").doc(customerId).set(newCustomer).then(() => {
                            console.log("Customer added successfully");
                        }).catch(error => console.error("Error adding document: ", error));

                    }

            
                    // No need to call saveCustomers() or filterAndRender() here.
                    // The `onSnapshot` listener in `loadCustomers` will handle updates automatically.


                    modal.style.display = 'none';

                    addCustomerForm.reset();

                    document.querySelector('#add-customer-modal h2').textContent = 'Add New Customer';

                    document.querySelector('#add-customer-form button').textContent = 'Add Customer';

                });

            

                dayFilters.addEventListener('click', (e) => {

                    if (e.target.classList.contains('day-filter')) {

                        document.querySelector('.day-filter.active').classList.remove('active');

                        e.target.classList.add('active');

                        updateSelectedDate(e.target.dataset.day);

                        filterAndRender();

                    }

                });

            

                searchInput.addEventListener('input', filterAndRender);

            

                customerGridBody.addEventListener('click', (e) => {

                    const row = e.target.closest('tr');

                    if (!row) return;

                    const weekId = getWeekId(new Date(currentDate));

            

                    // Handle Pay button click

                    if (e.target.classList.contains('pay-btn')) {

                        const customerId = row.querySelector('.customer-id').textContent;

                        const customerIndex = customers.findIndex(c => c.id === customerId);

                        if (customerIndex !== -1) {

                            const amountPaidInput = row.querySelector('.amount-paid-input');

                            const paymentModeSelect = row.querySelector('.payment-mode-select');

                            const paidAmount = parseInt(amountPaidInput.value);

            

                            if (isNaN(paidAmount) || paidAmount < 1) {

                                alert('Invalid amount. Please enter a number greater than 0.');

                                return;

                            }

            

                            const customerToUpdate = customers[customerIndex];
                            const newBalance = customerToUpdate.balanceAmount - paidAmount;
                            const newPaymentHistory = customerToUpdate.paymentHistory || {};
                            newPaymentHistory[weekId] = {
                                amount: paidAmount,
                                mode: paymentModeSelect.value,
                                status: 'Paid',
                                paymentDate: new Date()
                            };

                            db.collection("customers").doc(customerId).update({
                                balanceAmount: newBalance,
                                paymentHistory: newPaymentHistory
                            }).catch(error => console.error("Error updating document: ", error));

                        }

                        return;

                    }

            

                    // Handle Edit button click (for payment)

                    if (e.target.classList.contains('edit-pay-btn')) {

                        const customerId = row.querySelector('.customer-id').textContent;

                        const customerIndex = customers.findIndex(c => c.id === customerId);

                        if (customerIndex !== -1) {

                            const customerToUpdate = customers[customerIndex];
                            const payment = customerToUpdate.paymentHistory[weekId];
                            if (payment) {
                                const newBalance = customerToUpdate.balanceAmount + payment.amount;
                                const newPaymentHistory = customerToUpdate.paymentHistory;
                                delete newPaymentHistory[weekId];

                                db.collection("customers").doc(customerId).update({
                                    balanceAmount: newBalance,
                                    paymentHistory: newPaymentHistory
                                }).catch(error => console.error("Error updating document: ", error));
                            }

                                            }

                                            return;

                                        }

                        

                                        // Handle click on Name cell to show customer details

                                        if (e.target.closest('.customer-name-cell')) {

                                            const customerId = row.dataset.customerId;

                                            const customer = customers.find(c => c.id === customerId);

                                            if (customer) {

                                                showCustomerDetailsModal(customer);

                                            }

                                            return;

                                        }

                        

                                        // Handle click on Balance Amount cell to show transaction history

                                        if (e.target.closest('.balance-amount-cell')) {

                                            const customerId = row.dataset.customerId;

                                            const customer = customers.find(c => c.id === customerId);

                                            if (customer) {

                                                showTransactionHistoryModal(customer);

                                            }

                                            return;

                                        }

                                    });

                        

                                    // Event listener for closing the transaction history modal

                                    transactionHistoryCloseBtn.addEventListener('click', () => {

                                        transactionHistoryModal.style.display = 'none';

                                    });

                        

                                    window.addEventListener('click', (e) => {

                                        if (e.target === modal) {

                                            modal.style.display = 'none';

                                        }

                                                                                if (e.target === transactionHistoryModal) {

                                                                                    transactionHistoryModal.style.display = 'none';

                                                                                }

                                                                            });

                                        

                                                                                        // Initial calls to load data and render UI

                                        

                                            

                                        

                                                                                        const downloadExcelBtn = document.getElementById('download-excel-btn');

                                        

                                            

                                        

                                                                                        downloadExcelBtn.addEventListener('click', () => {

                                        

                                                                                            const activeDayFilterButton = document.querySelector('.day-filter.active');

                                        

                                                                                            const dayFilter = activeDayFilterButton ? activeDayFilterButton.dataset.day : 'All';

                                        

                                                                                            const weekId = getWeekId(new Date(currentDate));

                                        

                                            

                                        

                                                                                            let dataToExport = [];

                                        

                                                                                            const headers = ["ID", "Name", "Balance Amount", "Amount Paid", "Payment Mode", "Payment Status"];

                                        

                                                                                            dataToExport.push(headers);

                                        

                                            

                                        

                                                                                            // Filter customers based on the selected day

                                        

                                                                                            let customersForReport = customers;

                                        

                                                                                            if (dayFilter !== 'All') {

                                        

                                                                                                customersForReport = customers.filter(c => c.day === dayFilter);

                                        

                                                                                            }

                                        

                                            

                                        

                                                                                            customersForReport.forEach(customer => {

                                        

                                                                                                const payment = customer.paymentHistory ? customer.paymentHistory[weekId] : null;

                                        

                                                                                                const paymentStatus = payment ? payment.status : 'Pending';

                                        

                                                                                                const amountPaid = payment ? payment.amount : 0;

                                        

                                                                                                const paymentMode = payment ? payment.mode : 'N/A';

                                        

                                            

                                        

                                                                                                dataToExport.push([

                                        

                                                                                                    customer.id,

                                        

                                                                                                    customer.name,

                                        

                                                                                                    `₹${customer.balanceAmount.toLocaleString('en-IN')}`,

                                        

                                                                                                    `₹${amountPaid.toLocaleString('en-IN')}`,

                                        

                                                                                                    paymentMode,

                                        

                                                                                                    paymentStatus

                                        

                                                                                                ]);

                                        

                                                                                            });

                                        

                                            

                                        

                                                                                            const ws = XLSX.utils.aoa_to_sheet(dataToExport);

                                        

                                                                                            const wb = XLSX.utils.book_new();

                                        

                                                                                            XLSX.utils.book_append_sheet(wb, ws, "Daily Report");

                                        

                                                                                            XLSX.writeFile(wb, `Daily_Report_${dayFilter}_${new Date().toLocaleDateString('en-US').replace(/\//g, '-')}.xlsx`);

                                        

                                                                                        });

                                        

                                            

                                        


                                                                                        const downloadBookBtn = document.getElementById('download-book-btn');
                                                                                        downloadBookBtn.addEventListener('click', () => {
                                                                                            const wb = XLSX.utils.book_new();
                                                                                            const customersData = [];
                                                                                        
                                                                                            // Generate the last 15 week IDs from the current date
                                                                                            const last15WeekIds = [];
                                                                                            for (let i = 0; i < 15; i++) {
                                                                                                const date = new Date(); // Use current date, not from UI
                                                                                                date.setDate(date.getDate() - (i * 7));
                                                                                                last15WeekIds.push(getWeekId(date));
                                                                                            }
                                                                                            last15WeekIds.sort((a, b) => new Date(b) - new Date(a)); // Sort to have the most recent week first
                                                                                        
                                                                                            const headers = ['ID', 'Name', 'Total Payable', 'Amount Paid', 'Balance Amount', ...last15WeekIds.map(weekId => {
                                                                                                const date = new Date(weekId);
                                                                                                const day = String(date.getDate()).padStart(2, '0');
                                                                                                const month = String(date.getMonth() + 1).padStart(2, '0'); // Month is 0-indexed
                                                                                                const year = date.getFullYear();
                                                                                                return `${day}/${month}/${year}`;
                                                                                            })];
                                                                                            
                                                                                            customersData.push(headers);
                                                                                        
                                                                                            customers.forEach(customer => {
                                                                                                const totalPaid = Object.values(customer.paymentHistory || {}).reduce((sum, payment) => sum + payment.amount, 0);
                                                                                                const row = [customer.id, customer.name, customer.totalPayableAmount, totalPaid, customer.balanceAmount];
                                                                                                last15WeekIds.forEach(weekId => {
                                                                                                    const payment = customer.paymentHistory ? customer.paymentHistory[weekId] : null;
                                                                                                    row.push(payment ? payment.amount : 0);
                                                                                                });
                                                                                                customersData.push(row);
                                                                                            });
                                                                                        
                                                                                            const ws = XLSX.utils.aoa_to_sheet(customersData);

                                                                                            // Freeze columns A to E
                                                                                            ws['!freeze'] = { xSplit: 5, ySplit: 0, topLeftCell: 'F1', activePane: 'topRight', state: 'frozen' };

                                                                                            // Add distinct colors to columns A to E
                                                                                            const colors = ['FFFF00', '00FF00', '00FFFF', 'FF00FF', 'FFA500'];
                                                                                            for (let i = 0; i < 5; i++) {
                                                                                                const cellRef = XLSX.utils.encode_cell({c: i, r: 0});
                                                                                                if (!ws[cellRef].s) ws[cellRef].s = {};
                                                                                                if (!ws[cellRef].s.fill) ws[cellRef].s.fill = {};
                                                                                                ws[cellRef].s.fill.fgColor = { rgb: colors[i] };
                                                                                            }

                                                                                            XLSX.utils.book_append_sheet(wb, ws, 'Book');
                                                                                            XLSX.writeFile(wb, 'Book.xlsx');
                                                                                        });

                                                                                        loadCustomers();
                                                                                        preselectCurrentDayFilter();
                                                                                        updateWeekRange();
                                        
                                                                                        // filterAndRender() is now called by loadCustomers()
                                                                                        // after the initial data is loaded from Firestore.

                                        

                                                                                    });

            