const express = require('express');
const mysql = require('mysql2');
const bodyParser = require('body-parser');
const exceljs = require('exceljs');
const path = require('path');

// Initialize express
const app = express();

// Create MySQL connection
const db = mysql.createConnection({
  host: 'localhost',
  user: 'root',
  password: 'password',
  database: 'finance_app'
});

// Connect to MySQL
db.connect((err) => {
  if (err) throw err;
  console.log('Connected to the database');
});

// Set up middleware
app.use(bodyParser.urlencoded({ extended: true }));
app.set('view engine', 'ejs');
app.use(express.static('public'));

// Home route
app.get('/', (req, res) => {
  db.query('SELECT * FROM transactions ORDER BY date DESC', (err, results) => {
    if (err) throw err;

    let balance = 0;
    results.forEach((transaction) => {
      if (transaction.type === 'salary') {
        balance += transaction.amount;
      } else if (transaction.type === 'expense') {
        balance -= transaction.amount;
      }
    });

    res.render('index', { transactions: results, balance });
  });
});

// Add transaction route
app.post('/add', (req, res) => {
  const { type, amount, description, date, month } = req.body;
  db.query('INSERT INTO transactions (type, amount, description, date, month) VALUES (?, ?, ?, ?, ?)',
    [type, amount, description, date, month], (err, result) => {
      if (err) throw err;
      res.redirect('/');
    });
});

// Export to Excel route
app.get('/export', (req, res) => {
  db.query('SELECT * FROM transactions', (err, results) => {
    if (err) throw err;

    const workbook = new exceljs.Workbook();
    const worksheet = workbook.addWorksheet('Transactions');
    worksheet.columns = [
      { header: 'ID', key: 'id' },
      { header: 'Type', key: 'type' },
      { header: 'Amount', key: 'amount' },
      { header: 'Description', key: 'description' },
      { header: 'Date', key: 'date' },
      { header: 'Month', key: 'month' }
    ];

    worksheet.addRows(results);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="transactions.xlsx"');
    workbook.xlsx.write(res).then(() => res.end());
  });
});

// Set the port and listen for requests
app.listen(3001, () => {
  console.log('Server is running on http://localhost:3001');
});
