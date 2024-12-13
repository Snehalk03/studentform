<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    $fullName = $_POST['fullName'];
    $city = $_POST['city'];
    $email = $_POST['email'];
    $phone = $_POST['phone'];
    $domain = $_POST['domain'];
    $experience = $_POST['experience'];
    $companyName = $_POST['companyName'] ?? '';
    $yearsOfExperience = $_POST['yearsOfExperience'] ?? '';
    $designation = $_POST['designation'] ?? '';
    $ctc = $_POST['ctc'] ?? '';
    $gender = $_POST['gender'] ?? '';

    // Load or create the Excel file
    $file = 'Student_Registrations.xlsx';
    if (file_exists($file)) {
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
        $sheet = $spreadsheet->getActiveSheet();
    } else {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Registrations');
        $sheet->fromArray([
            'Full Name', 'City', 'Email', 'Phone', 'Domain', 'Experience',
            'Company Name', 'Years of Experience', 'Designation', 'CTC', 'Gender'
        ], NULL, 'A1');
    }

    // Add new data to the next available row
    $row = $sheet->getHighestRow() + 1;
    $sheet->fromArray([
        $fullName, $city, $email, $phone, $domain, $experience,
        $companyName, $yearsOfExperience, $designation, $ctc, $gender
    ], NULL, "A$row");

    // Save the updated Excel file
    $writer = new Xlsx($spreadsheet);
    $writer->save($file);

    echo "Registration successful. Data has been saved.";
} else {
    echo "Invalid request method.";
}

