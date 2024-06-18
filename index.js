
const XLSX = require('xlsx');
const fs = require('fs');


const workbook = XLSX.readFile('Employee_Data_Timesheet_May.xlsx');

const sheet_name_list = workbook.SheetNames;


const worksheet = workbook.Sheets[sheet_name_list[0]];
const worksheet1 = workbook.Sheets[sheet_name_list[1]];

const jsonData = XLSX.utils.sheet_to_json(worksheet);
const jsonData1 = XLSX.utils.sheet_to_json(worksheet1);




const parseEmployeeData = (data,data1) => {
  return data.map(row => {
    
   
    
      timesheet = data1.filter((x)=>{
        if(x.Employee===row.Employee){
            return x
        }
           
        
      })

    
    
    return {
      name: row['Employee Name'],
      Employee_ID:row['Employee'],
      salary: row['Salary'],
      employeeType: row['Employee Type'],
      timesheet
    };
  });
};



const employees = parseEmployeeData(jsonData,jsonData1);




// Function to calculate LOP
const calculateLOP = (timesheet) => {
  let lateOrEarlyCount = 0;


   timesheet.map((val)=>{

    if(val["Hours_Worked"]<9){
        lateOrEarlyCount++
    }

   })
   
  return Math.floor(lateOrEarlyCount / 3) * 0.5;
};



// Function to calculate salary
const calculateSalary = (employee) => {
  const baseSalary = employee.salary;
  if (employee.employeeType === 'Management') {
    return baseSalary;
  }

  const lop = calculateLOP(employee.timesheet);
  
  
  return baseSalary - (baseSalary / 30) * lop;
};




// Calculate final salary for each employee
employees.forEach(employee => {
  employee.finalSalary = calculateSalary(employee);
 


});





fs.writeFileSync('payroll.json', JSON.stringify(employees, null, 2));
console.log('Payroll calculated and saved to payroll.json');

let emp=employees.map((row)=>{
    return{
        name: row['name'],
        Employee_ID:row['Employee_ID'],
        salary: row['salary'],
        employeeType: row['employeeType'],
        Final_Salary: row['finalSalary']

    }
})

console.log(emp)

const writeJSONToExcel = (jsonData, outputFile) => {
    
    const worksheet = XLSX.utils.json_to_sheet(jsonData);
  
    
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
  
    
    XLSX.writeFile(workbook, outputFile);
    console.log(`Data written to ${outputFile}`);
  };
  
 
  const outputFile = 'Employee_Data.xlsx';
  
  // Write JSON data to Excel
  writeJSONToExcel(emp, outputFile);


