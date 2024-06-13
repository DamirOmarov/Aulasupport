const express = require('express');
const mongoose = require('mongoose');
const connectDB = require('./db'); // Импортируем функцию подключения к базе данных
const XLSX = require('xlsx'); // Импортируем библиотеку xlsx

const app = express();

// Подключение к MongoDB
connectDB();

// Схема и модель для коллекции ksk
const kskSchema = new mongoose.Schema({
    name: String,
    confirmed: Boolean
}, { collection: 'ksk' });

const Ksk = mongoose.model('Ksk', kskSchema);

// Схема и модель для коллекции requests
const requestSchema = new mongoose.Schema({
    company: {
        id: String,
        name: String
    },
    request_type: {
        id: String,
        name: String
    },
    created_at: { type: Date, default: Date.now },
    updated_at: { type: Date, default: Date.now },
    is_overdue: { type: Boolean, default: false }
}, { collection: 'requests' });

const Request = mongoose.model('Request', requestSchema);

app.get('/', async (req, res) => {
    try {
        const { startDate, endDate } = req.query;
        const dateFilter = {};

        if (startDate && endDate) {
            const start = new Date(`${startDate}T00:00:00.000Z`);
            const end = new Date(`${endDate}T23:59:59.999Z`);
            dateFilter.created_at = { $gte: start, $lte: end };
        } else if (startDate) {
            const start = new Date(`${startDate}T00:00:00.000Z`);
            dateFilter.created_at = { $gte: start };
        } else if (endDate) {
            const end = new Date(`${endDate}T23:59:59.999Z`);
            dateFilter.created_at = { $lte: end };
        }

        console.log('Date filter:', dateFilter);

        const confirmedData = await Ksk.find({ confirmed: true });
        const data = await Promise.all(confirmedData.map(async item => {
            const requestCount = await Request.countDocuments({ 'company.id': item._id.toString(), ...dateFilter });
            console.log(`Request count for company ${item.name}:`, requestCount);
            const specificRequestCount1 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda330796c90b114f5565', ...dateFilter });
            const specificRequestCount2 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda320796c90b114f555f', ...dateFilter });
            const specificRequestCount3 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda320796c90b114f5560', ...dateFilter });

            const specificRequestCount4NotOverdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda320796c90b114f5560',
                is_overdue: false,
                ...dateFilter
            });
            const specificRequestCount4Overdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda320796c90b114f5560',
                is_overdue: true,
                ...dateFilter
            });

            const specificRequestCount5NotOverdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda330796c90b114f5561',
                is_overdue: false,
                ...dateFilter
            });
            const specificRequestCount5Overdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda330796c90b114f5561',
                is_overdue: true,
                ...dateFilter
            });

            console.log(`Not overdue for company ${item.name} in "Назначено":`, specificRequestCount4NotOverdue);
            console.log(`Overdue for company ${item.name} in "Назначено":`, specificRequestCount4Overdue);
            console.log(`Not overdue for company ${item.name} in "В работе":`, specificRequestCount5NotOverdue);
            console.log(`Overdue for company ${item.name} in "В работе":`, specificRequestCount5Overdue);

            const specificRequestCount6 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda330796c90b114f5564', ...dateFilter });
            const specificRequestCount7 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '6373b7da9b87ee6862f76ddf', ...dateFilter });

            return {
                name: item.name,
                requestCount,
                specificRequestCount1,
                specificRequestCount2,
                specificRequestCount3,
                specificRequestCount4NotOverdue,
                specificRequestCount4Overdue,
                specificRequestCount5NotOverdue,
                specificRequestCount5Overdue,
                specificRequestCount6,
                specificRequestCount7
            };
        }));

        console.log('Data:', data);

        let responseHtml = `
        <html>
            <head>
                <title>KS List</title>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        background-color: #f4f4f4;
                        margin: 0;
                        padding: 20px;
                    }
                    h1 {
                        text-align: center;
                        color: #333;
                    }
                    form {
                        margin-bottom: 20px;
                        text-align: center;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        box-shadow: 0 2px 3px rgba(0,0,0,0.1);
                    }
                    th, td {
                        padding: 12px;
                        text-align: left;
                        border-bottom: 1px solid #ddd;
                    }
                    th {
                        background-color: #f8f8f8;
                        cursor: pointer;
                    }
                    tr:nth-child(even) {
                        background-color: #f9f9f9;
                    }
                    tr:hover {
                        background-color: #f1f1f1;
                    }
                    button {
                        padding: 10px 20px;
                        background-color: #4CAF50;
                        color: white;
                        border: none;
                        cursor: pointer;
                        font-size: 16px;
                    }
                    button:hover {
                        background-color: #45a049;
                    }
                    input[type="date"] {
                        padding: 8px;
                        font-size: 16px;
                    }
                    label {
                        margin: 0 10px;
                        font-size: 16px;
                    }
                    .overdue {
                        color: red;
                    }
                </style>
            </head>
            <body>
                <h1>Количество заявок по Управляющим компаниям</h1>
                <form method="GET" action="/">
                    <label for="startDate">С:</label>
                    <input type="date" id="startDate" name="startDate" value="${startDate || ''}">
                    <label for="endDate">По:</label>
                    <input type="date" id="endDate" name="endDate" value="${endDate || ''}">
                    <button type="submit">Filter</button>
                    <button type="button" onclick="exportToExcel()">Export to Excel</button>
                </form>
                <table border="1" id="kskTable">
                    <thead>
                        <tr>
                            <th onclick="sortTable(0)">No</th>
                            <th onclick="sortTable(1)">Name</th>
                            <th onclick="sortTable(2)">Общее</th>
                            <th onclick="sortTable(3)">Исполнено</th>
                            <th onclick="sortTable(4)">Новая</th>
                            <th onclick="sortTable(5)">Назначено</th>
                            <th onclick="sortTable(6)">В работе</th>
                            <th onclick="sortTable(7)">Отклонено</th>
                            <th onclick="sortTable(8)">Отменено</th>
                            <th onclick="sortTable(9)">Не исполнено</th>
                        </tr>
                    </thead>
                    <tbody>`;

        data.forEach((item, index) => {
            responseHtml += `
                        <tr>
                            <td>${index + 1}</td>
                            <td>${item.name}</td>
                            <td>${item.requestCount}</td>
                            <td>${item.specificRequestCount1}</td>
                            <td>${item.specificRequestCount2}</td>
                            <td>
                                <span>${item.specificRequestCount4NotOverdue}</span>
                                <span class="overdue">${item.specificRequestCount4Overdue}</span>
                            </td>
                            <td>
                                <span>${item.specificRequestCount5NotOverdue}</span>
                                <span class="overdue">${item.specificRequestCount5Overdue}</span>
                            </td>
                            <td>${item.specificRequestCount3}</td>
                            <td>${item.specificRequestCount6}</td>
                            <td>${item.specificRequestCount7}</td>
                        </tr>`;
        });

        responseHtml += `
                    </tbody>
                </table>
                <script>
                    function sortTable(n) {
                        var table, rows, switching, i, x, y, shouldSwitch, dir, switchcount = 0;
                        table = document.getElementById("kskTable");
                        switching = true;
                        dir = "asc"; 
                        while (switching) {
                            switching = false;
                            rows = table.rows;
                            for (i = 1; i < (rows.length - 1); i++) {
                                shouldSwitch = false;
                                x = rows[i].getElementsByTagName("TD")[n];
                                y = rows[i + 1].getElementsByTagName("TD")[n];
                                if (dir == "asc") {
                                    if (n === 0 || n === 2 || n === 3 || n === 4 || n === 5 || n === 6 || n === 7 || n === 8 || n === 9) { 
                                        if (parseInt(x.innerHTML) > parseInt(y.innerHTML)) {
                                            shouldSwitch = true;
                                            break;
                                        }
                                    } else {
                                        if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                                            shouldSwitch = true;
                                            break;
                                        }
                                    }
                                } else if (dir == "desc") {
                                    if (n === 0 || n === 2 || n === 3 || n === 4 || n === 5 || n === 6 || n === 7 || n === 8 || n === 9) { 
                                        if (parseInt(x.innerHTML) < parseInt(y.innerHTML)) {
                                            shouldSwitch = true;
                                            break;
                                        }
                                    } else {
                                        if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                                            shouldSwitch = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            if (shouldSwitch) {
                                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                                switching = true;
                                switchcount ++;
                            } else {
                                if (switchcount == 0 && dir == "asc") {
                                    dir = "desc";
                                    switching = true;
                                }
                            }
                        }
                    }

                    function exportToExcel() {
                        window.location.href = '/export?startDate=' + document.getElementById('startDate').value + '&endDate=' + document.getElementById('endDate').value;
                    }
                </script>
            </body>
        </html>`;

        res.send(responseHtml);
    } catch (err) {
        console.error(err);
        res.status(500).send(err.message);
    }
});

app.get('/export', async (req, res) => {
    try {
        const { startDate, endDate } = req.query;
        const dateFilter = {};

        if (startDate && endDate) {
            const start = new Date(`${startDate}T00:00:00.000Z`);
            const end = new Date(`${endDate}T23:59:59.999Z`);
            dateFilter.created_at = { $gte: start, $lte: end };
        } else if (startDate) {
            const start = new Date(`${startDate}T00:00:00.000Z`);
            dateFilter.created_at = { $gte: start };
        } else if (endDate) {
            const end = new Date(`${endDate}T23:59:59.999Z`);
            dateFilter.created_at = { $lte: end };
        }

        const confirmedData = await Ksk.find({ confirmed: true });
        const data = await Promise.all(confirmedData.map(async item => {
            const requestCount = await Request.countDocuments({ 'company.id': item._id.toString(), ...dateFilter });
            const specificRequestCount1 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda330796c90b114f5565', ...dateFilter });
            const specificRequestCount2 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda320796c90b114f555f', ...dateFilter });
            const specificRequestCount3 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda320796c90b114f5560', ...dateFilter });

            const specificRequestCount4NotOverdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda320796c90b114f5560',
                is_overdue: false,
                ...dateFilter
            });
            const specificRequestCount4Overdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda320796c90b114f5560',
                is_overdue: true,
                ...dateFilter
            });

            const specificRequestCount5NotOverdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda330796c90b114f5561',
                is_overdue: false,
                ...dateFilter
            });
            const specificRequestCount5Overdue = await Request.countDocuments({
                'company.id': item._id.toString(),
                'request_type.id': '5f4cda330796c90b114f5561',
                is_overdue: true,
                ...dateFilter
            });

            const specificRequestCount6 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '5f4cda330796c90b114f5564', ...dateFilter });
            const specificRequestCount7 = await Request.countDocuments({ 'company.id': item._id.toString(), 'request_type.id': '6373b7da9b87ee6862f76ddf', ...dateFilter });

            return {
                name: item.name,
                Общее: requestCount,
                Исполнено: specificRequestCount1,
                Новая: specificRequestCount2,
                Назначено: specificRequestCount3,
                'Назначено (не просрочено)': specificRequestCount4NotOverdue,
                'Назначено (просрочено)': specificRequestCount4Overdue,
                'В работе (не просрочено)': specificRequestCount5NotOverdue,
                'В работе (просрочено)': specificRequestCount5Overdue,
                Отклонено: specificRequestCount6,
                'Не исполнено': specificRequestCount7
            };
        }));

        // Добавляем заголовки колонок
        const headers = [
            "name",
            "Общее",
            "Исполнено",
            "Новая",
            "Назначено",
            "Назначено (не просрочено)",
            "Назначено (просрочено)",
            "В работе (не просрочено)",
            "В работе (просрочено)",
            "Отклонено",
            "Не исполнено"
        ];

        const worksheet = XLSX.utils.json_to_sheet(data, { header: headers });
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');

        const excelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
        const fileName = `export_${Date.now()}.xlsx`;

        res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(excelBuffer);
    } catch (err) {
        console.error(err);
        res.status(500).send(err.message);
    }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
