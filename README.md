# Easy Report

**Easy Report** is a Python desktop application for generating monthly student work reports. It helps teachers create reports and allows managers to monitor student progress efficiently.

---

## Features

- Select Excel files with student homework and student information.
- Generate a **Manager Excel Report** summarizing all student performance.
- Generate **Student Word Reports** with individual topics, grades, and footnotes.
- Supports multiple languages: English, Persian, Hindi.
- User-friendly GUI built with **Tkinter**.
- Progress bar to track report generation.
- Support page with contact information.
- Flexible save locations for reports.

---

## Requirements

- Python 3.x
- Libraries:

```bash
pip install pandas numpy openpyxl python-docx
```

## Installation

1. **Clone the repository**:

```bash
git clone https://github.com/yourusername/EasyReport.git
```

2. **Navigate into the project folder:**:

```bash
cd EasyReport/projectFile
```

3. **Run the program:**:

```bash
python EasyReprt.py
```

## Usage

1. **Select Files in the GUI**:

   - **Student Homework Excel File** – Excel file with homework data entered by teachers.  
   - **Student Information Excel File** – Excel file containing each student's details such as name, email, and invoice number.  
   - **Student Word Report Template** – Word (.docx) template with placeholders.  
   - **Save Folder** – Location to save the generated reports.

2. **Choose Action Type**:

   - `Manager Excel Report` – Generates a report summarizing student grades for managers.  
   - `Students Word Report` – Generates personalized Word reports for each student.  
   - `All` – Generates both manager and student reports.

3. **Select Month and Year** for filtering the data.

4. Click **Start** to generate reports. The progress bar will indicate the status.

---

## Output

- **Manager Excel Report**: `Manager_Report_{Month}_{Year}.xlsx`  
- **Student Word Reports Folder**: `Student_Word_Report_{Month}_{Year}` containing individual Word files for each student.

---

## Word Template Placeholders

| Placeholder      | Replaced With                     |
|-----------------|----------------------------------|
| `«Student_Name»` | Student's full name               |
| `«Foot_Notes»`   | Student-specific footnotes        |
| `«Topic_Grades»` | Table of topics and average grades |

**Note:** Placeholders must exist in the Word template for proper replacement during report generation.

---

## File Structure

- `EasyReprt.py` – Main Python script with GUI and report generation logic.  
- `Student Word Template.docx` – Template file for generating student Word reports.  
- `Excel files` – Teacher homework data and student info files provided as input.

---

## Support

For any issues regarding Excel or Word report generation:

- 📧 Email: `parsa.asil@outlook.com`  
- ⏰ Working Hours: Mon-Fri, 9:00 AM - 6:00 PM  

---

## License

This project is open-source. You can modify and use it for educational or personal purposes.
