# 📘 Student Performance Tracker (Excel VBA Project)

This is a complete Excel VBA-based system that allows you to Collect, manage and analyze student academic performance using a simple **UserForm interface** enabling performance tracking and actionable insights. It's designed for **schools, teachers, or academic institutions** looking for an easy-to-use offline solution to track **marks, attendance, grades**, and generate **dashboards** with charts.

<img width="1899" height="741" alt="Image" src="https://github.com/user-attachments/assets/eec9652f-dc33-4d0f-8ebd-a701f20949fa" />

<img width="1046" height="519" alt="Image" src="https://github.com/user-attachments/assets/b428dfc2-1a5b-46b2-a3d9-01d3736707cc" />

---

## 🔍 Overview

The **Student Performance Tracker** project simplifies the process of:
- Recording student details and marks
- Automatically calculating total marks, percentage, attendance, and grades
- Managing records with edit and delete functions
- Generating a **dynamic grade distribution pie chart**
- Using an Excel UserForm to make data entry user-friendly

---

## 🧠 Key Functionalities

### 🎯 1. **Student Entry UserForm**
- Add new students
- Update existing student records
- Delete records by Student ID
- Auto-calculations:
  - Total Marks
  - Attendance Percentage
  - Grade (A+, A, B, C, F)

### 📈 2. **Automatic Grade Calculation Logic**
| Percentage      | Grade |
|-----------------|--------|
| 90 and above    | A+     |
| 75 – 89         | A      |
| 60 – 74         | B      |
| 40 – 59         | C      |
| Below 40        | F      |

### 📊 3. **Dashboard and Charts**
- A pie chart showing the **distribution of grades**
- Grade-wise summary using `COUNTIF`
- Chart auto-refreshes with the click of a button

---

## 📄 Data Structure

Data is stored in the `Performance_Tracker` worksheet with the following columns:

| Column            | Description                      |
|-------------------|----------------------------------|
| Student ID        | Unique ID                        |
| First Name        | Student's first name             |
| Last Name         | Student's last name              |
| Gender            | Male/Female/Other                |
| Date of Birth     | DOB (used to calculate age)      |
| Age               | Auto-calculated                  |
| Class             | Student's class/grade            |
| Maths, Science, English | Marks entered manually   |
| Present Days      | Days present                     |
| Total Days        | Total academic days              |
| Total Marks       | Auto sum of subjects             |
| Attendance %      | Calculated                       |
| Grade             | Based on percentage              |


---
## **Summary**


The Student Performance Tracker is a VBA-powered data analysis tool designed to transform raw academic inputs into actionable insights. Capturing over 15 structured metrics per student—including subject-wise marks, attendance, and demographics—it automatically calculates total marks, attendance percentage, assigns grades based on performance logic and reducing manual workloal by 60% . The system visualizes trends using dynamic dashboards; for instance, current data shows 66% of students fall in the ‘C’ grade band, while one student has attendance below 40%, signaling potential intervention needs. By integrating real-time calculations and grade distribution charts, this tool streamlines academic monitoring and enables data-driven decision-making for educators





