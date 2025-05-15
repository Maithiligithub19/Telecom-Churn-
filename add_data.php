<?php
// require_once 'PHPExcel/IOFactory.php';
require 'PHPExcel/Classes/PHPExcel.php';

if ($_SERVER["REQUEST_METHOD"] == "POST") {
    $customerID = $_POST["customerID"];
    $gender = $_POST["gender"];
    $SeniorCitizen = $_POST["SeniorCitizen"];
    $Partner = $_POST["Partner"];
    $Dependents = $_POST["Dependents"];
    $tenure = $_POST["tenure"];
    $PhoneService = $_POST["PhoneService"];
    $MultipleLines = $_POST["MultipleLines"];
    $InternetService = $_POST["InternetService"];
    $OnlineSecurity = $_POST["OnlineSecurity"];
    $OnlineBackup = $_POST["OnlineBackup"];
    $DeviceProtection = $_POST["DeviceProtection"];
    $TechSupport = $_POST["TechSupport"];
    $StreamingTV = $_POST["StreamingTV"];
    $StreamingMovies = $_POST["StreamingMovies"];
    $Contract = $_POST["Contract"];
    $PaperlessBilling = $_POST["PaperlessBilling"];
    $PaymentMethod = $_POST["PaymentMethod"];
    $MonthlyCharges = $_POST["MonthlyCharges"];
    $TotalCharges = $_POST["TotalCharges"];
    $Churn = $_POST["Churn"];

    $excelFilePath = 'E:\\powerbi\\Telco-Customer-Churn.csv';

    // Load the Excel file
    $objPHPExcel = PHPExcel_IOFactory::load($excelFilePath);

    // Select the first worksheet
    $worksheet = $objPHPExcel->getActiveSheet();

    // Find the last row in the worksheet
    $lastRow = $worksheet->getHighestRow() + 1;

    // Add the user's data to the Excel file
    $worksheet->setCellValue('customerID' . $lastRow, $customerID);
    $worksheet->setCellValue('gender' . $lastRow, $gender);
    $worksheet->setCellValue('SeniorCitizen' . $lastRow, $SeniorCitizen);
    $worksheet->setCellValue('Partner' . $lastRow, $Partner);
    $worksheet->setCellValue('Dependents' . $lastRow, $Dependents);
    $worksheet->setCellValue('tenure' . $lastRow, $tenure);
    $worksheet->setCellValue('PhoneService' . $lastRow, $PhoneService);
    $worksheet->setCellValue('MultipleLines' . $lastRow, $MultipleLines);
    $worksheet->setCellValue('InternetService' . $lastRow, $InternetService);
    $worksheet->setCellValue('OnlineSecurity' . $lastRow, $OnlineSecurity);
    $worksheet->setCellValue('OnlineBackup' . $lastRow, $OnlineBackup);
    $worksheet->setCellValue('DeviceProtection' . $lastRow, $DeviceProtection);
    $worksheet->setCellValue('TechSupport' . $lastRow, $TechSupport);
    $worksheet->setCellValue('StreamingTV' . $lastRow, $StreamingTV);
    $worksheet->setCellValue('StreamingMovies' . $lastRow, $StreamingMovies);
    $worksheet->setCellValue('Contract' . $lastRow, $Contract);
    $worksheet->setCellValue('PaperlessBilling' . $lastRow, $PaperlessBilling);
    $worksheet->setCellValue('PaymentMethod' . $lastRow, $PaymentMethod);
    $worksheet->setCellValue('MonthlyCharges' . $lastRow, $MonthlyCharges);
    $worksheet->setCellValue('TotalCharges' . $lastRow, $TotalCharges);
    $worksheet->setCellValue('Churn' . $lastRow, $Churn);

    // Save the changes
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2019');
    $objWriter->save($excelFilePath);

    header("Location: index.html"); // Redirect to the form page
}
?>