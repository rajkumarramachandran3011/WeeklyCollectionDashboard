document.addEventListener('DOMContentLoaded', () => {
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

    let currentDate = new Date();
    let customers = [];

    // Load customers from Local Storage
    const loadCustomers = () => {
        const storedCustomers = localStorage.getItem('customers');
        if (storedCustomers) {
            customers = JSON.parse(storedCustomers);
        }
    };

    // Save customers to Local Storage
    const saveCustomers = () => {
        localStorage.setItem('customers', JSON.stringify(customers));
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
        const firstDay = new Date(date.setDate(date.getDate() - date.getDay()));
        return firstDay.toISOString().split('T')[0];
    }

    const renderGrid = (customerData) => {
        const weekId = getWeekId(new Date(currentDate));
        customerGridBody.innerHTML = '';
        for (const customer of customerData) {
            const payment = customer.paymentHistory ? customer.paymentHistory[weekId] : null;
            const paymentStatus = payment ? payment.status : 'Pending';
            const lastPaidAmount = payment ? payment.amount : 0;
            const paymentMode = payment ? payment.mode : 'Cash';

                        const row = document.createElement('tr');

                        row.innerHTML = `

                            <td data-label="Name">

                                <div class="customer-name">${customer.name}</div>

                                <div class="customer-id">${customer.id}</div>

                            </td>

                            <td data-label="Day">

                                <div>${customer.day}</div>

                                <div class="customer-address">${customer.address}</div>

                            </td>

                            <td data-label="Balance Amount">

                                <div class="balance-amount">₹${customer.balanceAmount.toLocaleString('en-IN')}</div>

                                <div class="total-payable">of ₹${customer.totalPayableAmount.toLocaleString('en-IN')}</div>

                            </td>

                            <td data-label="Amount Paid (Current)"><input type="number" class="amount-paid-input" value="${paymentStatus === 'Paid' ? lastPaidAmount : ''}" ${paymentStatus === 'Paid' ? 'disabled' : ''} min="1" max="99999"></td>

             

                            <td data-label="Payment Mode">

                                <select class="payment-mode-select" ${paymentStatus === 'Paid' ? 'disabled' : ''}>

                                    <option value="Cash" ${paymentMode === 'Cash' ? 'selected' : ''}>Cash</option>

                                    <option value="UPI" ${paymentMode === 'UPI' ? 'selected' : ''}>UPI</option>

                                </select>

                            </td>

                            <td data-label="Actions">

                                ${paymentStatus === 'Paid' ? '<button class="edit-pay-btn"><i class="fas fa-edit"></i> Edit</button>' : '<button class="pay-btn"><i class="fas fa-money-bill-wave"></i> Pay</button>'}

                            </td>

                            <td data-label="Payment Status">

                                ${paymentStatus === 'Paid' ? `<span class="paid-status">Paid</span><div class="paid-date">${new Date(payment.paymentDate).toLocaleString()}</div>` : '<span class="not-paid-status">Not Paid</span>'}

                            </td>

                        `;

                        customerGridBody.appendChild(row);

                    }

                };

            

                let editingCustomerIndex = null;

            

            

                const filterAndRender = () => {

                    const dayFilter = document.querySelector('.day-filter.active').dataset.day;

                    const searchTerm = searchInput.value.toLowerCase();

            

                    let filteredCustomers = customers;

            

                    if (dayFilter !== 'All') {

                        filteredCustomers = filteredCustomers.filter(c => c.day === dayFilter);

                    }

            

                    filteredCustomers = filteredCustomers.filter(c => {

                        const accountOpeningDate = new Date(c.accountOpeningDate);

                        const oneDayAgo = new Date(currentDate);

                        oneDayAgo.setDate(currentDate.getDate() - 1); // Set to 1 day before currentDate

                        return accountOpeningDate < oneDayAgo;

                    });

            

                    if (searchTerm) {

                        filteredCustomers = filteredCustomers.filter(c => 

                            c.name.toLowerCase().includes(searchTerm) || 

                            c.phone.toLowerCase().includes(searchTerm) ||

                            (c.address && c.address.toLowerCase().includes(searchTerm))

                        );

                    }

            

                    renderGrid(filteredCustomers);

                };

            

                const deleteCustomerBtn = document.getElementById('delete-customer-btn');

            

                addCustomerBtn.addEventListener('click', () => {

                    editingCustomerIndex = null;

                    document.querySelector('#add-customer-modal h2').textContent = 'Add New Customer';

                    document.querySelector('#add-customer-form button').textContent = 'Add Customer';

                    addCustomerForm.reset();

                    document.getElementById('customer-id').value = getNextCustomerId();

                    deleteCustomerBtn.style.display = 'none'; // Hide delete button when adding new customer

                    modal.style.display = 'block';

                });

            

                closeBtn.addEventListener('click', () => {

                    modal.style.display = 'none';

                });

            

                deleteCustomerBtn.addEventListener('click', () => {

                    if (editingCustomerIndex !== null) {

                        const confirmation = confirm('Are you sure you want to delete this customer? This action cannot be undone.');

                        if (confirmation) {

                            customers.splice(editingCustomerIndex, 1);

                            saveCustomers();

                            editingCustomerIndex = null;

                            filterAndRender();

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

                        id: document.getElementById('customer-id').value,

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

                        customers[editingCustomerIndex] = newCustomer;

                        editingCustomerIndex = null;

                    } else {

                        customers.push(newCustomer);

                    }

            

                    saveCustomers();

                    filterAndRender();

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

            

                            customers[customerIndex].balanceAmount -= paidAmount; // Subtract from balance

                            if (!customers[customerIndex].paymentHistory) {

                                customers[customerIndex].paymentHistory = {};

                            }

                            customers[customerIndex].paymentHistory[weekId] = {

                                amount: paidAmount,

                                mode: paymentModeSelect.value,

                                status: 'Paid',

                                paymentDate: new Date()

                            };

                            saveCustomers(); // Save updated customers to local storage

                            filterAndRender();

                        }

                        return;

                    }

            

                    // Handle Edit button click (for payment)

                    if (e.target.classList.contains('edit-pay-btn')) {

                        const customerId = row.querySelector('.customer-id').textContent;

                        const customerIndex = customers.findIndex(c => c.id === customerId);

                        if (customerIndex !== -1) {

                            const payment = customers[customerIndex].paymentHistory[weekId];

                            customers[customerIndex].balanceAmount += payment.amount;

                            delete customers[customerIndex].paymentHistory[weekId];

                            saveCustomers(); // Save updated customers to local storage

                            filterAndRender();

                        }

                        return;

                    }

            

                    // Handle click on first 2 columns to open popup (Name, Day)

                    const clickedCell = e.target.closest('td');

                    if (clickedCell && clickedCell.cellIndex >= 0 && clickedCell.cellIndex <= 2) { // Name, Day, Balance

                        const customerId = row.querySelector('.customer-id').textContent;

                        const customerIndex = customers.findIndex(c => c.id === customerId);

                        const customer = customers[customerIndex];

            

                        document.querySelector('#add-customer-modal h2').textContent = 'Edit Customer';

                        document.querySelector('#add-customer-form button[type="submit"]').textContent = 'Update Customer';

            

                        document.getElementById('customer-id').value = customer.id;

                        document.getElementById('name').value = customer.name;

                        document.getElementById('phone').value = customer.phone;

                        document.getElementById('address').value = customer.address;

                        document.getElementById('day').value = customer.day;

            

                        document.getElementById('total-payable-amount').value = customer.totalPayableAmount;

                        document.getElementById('number-of-installments').value = customer.numberOfInstallments;

                        // The balance-amount input is removed, so we don't set its value here.

                        document.getElementById('amount-paid-input').value = customer.amountPaid;

                        document.getElementById('account-opening-date').value = customer.accountOpeningDate;

            

                        deleteCustomerBtn.style.display = 'block'; // Show delete button when editing

                        editingCustomerIndex = customerIndex;

                        modal.style.display = 'block';

                    }

                });

            

                const updateWeekRange = () => {

                    const firstDay = new Date(new Date(currentDate).setDate(currentDate.getDate() - currentDate.getDay()));

                    const lastDay = new Date(new Date(currentDate).setDate(currentDate.getDate() - currentDate.getDay() + 6));

                    const options = { month: 'short', day: 'numeric', year: 'numeric' };

                    weekRangeEl.textContent = `${firstDay.toLocaleDateString('en-US', options)} - ${lastDay.toLocaleDateString('en-US', options)}`;

                    updateSelectedDate(document.querySelector('.day-filter.active').dataset.day);

                    filterAndRender();

                };

            

                const updateSelectedDate = (day) => {

                    if (day === 'All') {

                        selectedDateDisplay.textContent = '';

                        return;

                    }

                    const dayIndex = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'].indexOf(day);

                    const firstDayOfWeek = new Date(new Date(currentDate).setDate(currentDate.getDate() - currentDate.getDay()));

                    const selectedDate = new Date(firstDayOfWeek.setDate(firstDayOfWeek.getDate() + dayIndex));

                    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };

                    selectedDateDisplay.textContent = selectedDate.toLocaleDateString('en-US', options);

                };

            

                prevWeekBtn.addEventListener('click', () => {

                    currentDate.setDate(currentDate.getDate() - 7);

                    updateWeekRange();

                });

            

                nextWeekBtn.addEventListener('click', () => {

                    currentDate.setDate(currentDate.getDate() + 7);

                    updateWeekRange();

                });

            

                const generateReportBtn = document.getElementById('generate-report-btn');

            

                generateReportBtn.addEventListener('click', () => {

                    const weekId = getWeekId(new Date(currentDate));

                    const reportData = [

                        ['ID', 'Name', 'Date', 'Total Payable Amount', 'Balance Amount', 'Amount Paid', 'Payment Mode', 'Payment Status']

                    ];

            

                    const dayFilter = document.querySelector('.day-filter.active').dataset.day;

                    let reportDate = new Date(currentDate);

                    if (dayFilter !== 'All') {

                        const dayIndex = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'].indexOf(dayFilter);

                        const firstDayOfWeek = new Date(new Date(currentDate).setDate(currentDate.getDate() - currentDate.getDay()));

                        reportDate = new Date(firstDayOfWeek.setDate(firstDayOfWeek.getDate() + dayIndex));

                    }

            

                    customers.forEach(customer => {

                        const payment = customer.paymentHistory ? customer.paymentHistory[weekId] : null;

                        const paymentStatus = payment ? payment.status : 'Pending';

                        const amountPaid = payment ? payment.amount : 0;

                        const paymentMode = payment ? payment.mode : 'N/A';

                        const reportRow = [

                            customer.id,

                            customer.name,

                            reportDate.toLocaleDateString(),

                            customer.totalPayableAmount,

                            customer.balanceAmount,

                            amountPaid,

                            paymentMode,

                            paymentStatus

                        ];

                        reportData.push(reportRow);

                    });

            

                    const ws = XLSX.utils.aoa_to_sheet(reportData);

                    const wb = XLSX.utils.book_new();

                    XLSX.utils.book_append_sheet(wb, ws, 'Weekly Report');

            

                    // Add styling

                    const headerStyle = {

                        font: {

                            bold: true,

                            color: { rgb: "FFFFFF" }

                        },

                        fill: {

                            fgColor: { rgb: "6a0dad" }

                        }

                    };

            

                    const headerRange = XLSX.utils.decode_range(ws['!ref']);

                    for (let C = headerRange.s.c; C <= headerRange.e.c; ++C) {

                        const cell = ws[XLSX.utils.encode_cell({ r: 0, c: C })];

                        cell.s = headerStyle;

                    }

            

                    XLSX.writeFile(wb, 'weekly_report.xlsx');

                });

            

                loadCustomers();

                updateWeekRange();

            });

            