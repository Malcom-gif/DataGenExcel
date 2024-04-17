const faker = require('faker');
const Excel = require('exceljs');

async function generateExcel() {
  const workbook = new Excel.Workbook();
  const sheet = workbook.addWorksheet('Data');

  // Define columns based on your sample
  sheet.columns = [
    { header: 'Gender', key: 'gender', width: 10 },
    { header: 'Age', key: 'age', width: 10 },
    { header: 'Income', key: 'income', width: 10 },
    { header: 'Loan Amount', key: 'loan_amount', width: 15 },
    { header: 'Employee Status', key: 'employee_status', width: 15 },
    { header: 'Loan Purpose', key: 'loan_purpose', width: 15 },
    { header: 'Collateral', key: 'collateral', width: 15 },
    { header: 'Marital Status', key: 'marital_status', width: 15 },
    { header: 'Account Number', key: 'account_number', width: 20 },
    { header: 'Credit Score', key: 'credit_score', width: 15 },
    { header: 'Interest Rate', key: 'interest_rate', width: 15 },
    { header: 'Debt to Income Ratio', key: 'debt_to_income_ratio', width: 20 },
    { header: 'Education', key: 'education', width: 15 },
    { header: 'Residential Status', key: 'residential_status', width: 20 },
    { header: 'Classification', key: 'classification', width: 15 },
  ];

  // Generate random data
  for (let i = 0; i < 5000; i++) {
    sheet.addRow({
      gender: faker.random.arrayElement(['Male', 'Female']),
      age: faker.datatype.number({ min: 18, max: 70 }),
      income: faker.datatype.number({ min: 300, max: 5000 }),
      loan_amount: faker.datatype.number({ min: 300, max: 10000 }),
      employee_status: faker.random.arrayElement(['Employed', 'Unemployed', 'Self-Employed']),
      loan_purpose: faker.random.arrayElement(['Personal', 'Business']),
      collateral: faker.random.arrayElement(['none', 'presented']),
      marital_status: faker.random.arrayElement(['Single', 'Married', 'Divorced']),
      account_number: faker.datatype.number({ min: 1e10, max: 1e11 - 1 }),
      credit_score: faker.datatype.number({ min: 300, max: 850 }),
      interest_rate: faker.datatype.float({ min: 0, max: 25 }).toFixed(2),
      debt_to_income_ratio: faker.datatype.float({ min: 0, max: 40 }).toFixed(2),
      education: faker.random.arrayElement(['Degree', 'Diploma', 'Certificate']),
      residential_status: faker.random.arrayElement(['Owned', 'Rented', 'Other', 'Parents']),
      classification: faker.random.arrayElement(['Good', 'Bad']),
    });
  }

  // Write to a file
  await workbook.xlsx.writeFile('RandomData.xlsx');
  console.log('Excel file created!');
}

generateExcel().catch(console.error);
