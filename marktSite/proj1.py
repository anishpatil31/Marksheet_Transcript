import os
import csv
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font
import sys
from .models import *
from django.shortcuts import render, redirect
from .forms import *
import openpyxl
from openpyxl.styles import Font


def projf(request):
    v = Marktsheet_data.objects.all()
    positive = 0
    negative = 0
    for p in v:
        positive = p.positive
        negative = p.negative
    answer_key = []
    output_path = r"marksheets//"
    output_path2 = r"marksheets"
    if os.path.exists(output_path2) == False:
        os.mkdir(r"marksheets")
    path = ""
    response = r"media/file/responses.csv"
    master = r"media/file/master_roll.csv"
    logo_path = r"media/image/Logo.jpeg"
    flag = "anish"

    with open(os.path.join(path, response), 'r') as responses:
        reader = csv.reader(responses)
        heading_of_columns = next(reader)
        for row in reader:
            if row[6] == "ANSWER":
                for i in range(7, 35):
                    answer_key.append(row[i])
                flag = row[6]
                break

        if flag != "ANSWER":
            if request.method == 'POST':
                form = Markform(request.POST, request.FILES)

                if form.is_valid():
                    form.save()
                    return redirect('Proj1')
            else:
                form = Markform()
            return render(request, 'marktSite/marktSite.html', {'form': form, 'check': 1})

    correct_marks = 0
    negative_marks = 0

    correct_marks = positive
    negative_marks = negative

    with open(os.path.join(path, master), 'r') as file:
        reader = csv.reader(file)
        heading_of_columns = next(reader)
        for row in reader:
            roll = row[0]
            name = row[1]
            file_path = output_path + '%s.xlsx' % roll
            wb = Workbook()
            wb_sheet = wb['Sheet']
            wb_sheet.title = 'quiz'
            sheet = wb.active
            img = openpyxl.drawing.image.Image(logo_path)
            img.anchor = 'A1'
            sheet.add_image(img)
            sheet.append([])
            sheet.append([])
            sheet.append([])
            sheet.append(["", "", "Mark Sheet"])
            heading = ["Name: ", name, " ", "Exam: ", "quiz"]
            row = ["Roll: ", roll]
            sheet.append(heading)
            sheet.append(row)
            found = 0

            with open(os.path.join(path, response), 'r') as f:
                read = csv.reader(f)
                for rows in read:
                    if rows[6] == roll:
                        found = 1
                        blank_row = [" "]
                        new_row = [" ", "Right", "Wrong", "Not Attempt", "Max"]
                        sheet.append(blank_row)
                        sheet.append(new_row)
                        right = 0
                        wrong = 0
                        not_attempted = 0
                        for i in range(0, len(answer_key)):
                            if rows[i + 7] == answer_key[i]:
                                right += 1
                            elif rows[i + 7] == "":
                                not_attempted += 1
                            else:
                                wrong += 1
                        total_questions = right + wrong + not_attempted
                        total_marks = total_questions * correct_marks
                        student_marks = correct_marks * right + negative_marks * wrong
                        print_student = str(student_marks) + "/" + str(total_marks)
                        no = ["No.", right, wrong, not_attempted, total_questions]
                        marking = ["Marking", correct_marks, negative_marks, "0"]
                        total_marks = ["Total", correct_marks * right, negative_marks * wrong, " ", print_student]
                        student_ans = ["Student Ans", "Correct Ans", " ", "Student Ans", "Correct Ans"]
                        sheet.append(no)
                        sheet.append(marking)
                        sheet.append(total_marks)
                        sheet.append(blank_row)
                        sheet.append(blank_row)
                        sheet.append(student_ans)
                        for i in range(0, 3):
                            new_row = [rows[7 + i], answer_key[i], " ", rows[32 + i], answer_key[25 + i]]
                            sheet.append(new_row)
                            z = 15 + i
                            s = "A" + str(z)
                            font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                        underline='none', strike=False, color='00A300')
                            if rows[7 + i] != answer_key[i]:
                                font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                            underline='none', strike=False, color='FF0000')
                            sheet[s].font = font

                            s = "D" + str(z)
                            font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                        underline='none', strike=False, color='00A300')
                            if rows[32 + i] != answer_key[25 + i]:
                                font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                            underline='none', strike=False, color='FF0000')
                            sheet[s].font = font

                        for i in range(3, 28):
                            new_row = [rows[7 + i], answer_key[i]]
                            sheet.append(new_row)
                            z = 15 + i
                            s = "A" + str(z)
                            font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                        underline='none', strike=False, color='00A300')
                            if rows[7 + i] != answer_key[i]:
                                font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                            underline='none', strike=False, color='FF0000')
                            sheet[s].font = font

                        font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None, underline='none',
                                    strike=False, color='000000')
                        sheet['B8'].font = font
                        sheet['C8'].font = font
                        sheet['D8'].font = font
                        sheet['E8'].font = font
                        sheet['A9'].font = font
                        sheet['A10'].font = font
                        sheet['A11'].font = font
                        font = Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                                    strike=False, color='FF0000')
                        sheet['C9'].font = font
                        sheet['C10'].font = font
                        sheet['C11'].font = font
                        # green
                        font = Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                                    strike=False, color='00A300')
                        sheet['B9'].font = font
                        sheet['B10'].font = font
                        sheet['B11'].font = font

                        # blue
                        font = Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                                    strike=False, color='0000FF')
                        sheet['E11'].font = font
                        for i in range(15, 43):
                            s = "B" + str(i)
                            sheet[s].font = font
                            if i < 18:
                                s = "E" + str(i)
                                sheet[s].font = font

                        font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None,
                                    underline='single', strike=False, color='000000')
                        sheet['C4'].font = font

                        # black
                        font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None, underline='none',
                                    strike=False, color='000000')
                        sheet['B5'].font = font
                        sheet['B6'].font = font
                        sheet['E5'].font = font
                        sheet['A14'].font = font
                        sheet['B14'].font = font
                        sheet['D14'].font = font
                        sheet['E14'].font = font

            if found == 0:
                blank_row = [" "]
                new_row = [" ", "Right", "Wrong", "Not Attempt", "Max"]
                sheet.append(blank_row)
                sheet.append(new_row)
                right = 0
                wrong = 0
                not_attempted = len(answer_key)
                no = ["No.", right, wrong, not_attempted, not_attempted]
                marking = ["Marking", correct_marks, negative_marks, "0"]
                total_marks = ["Total", correct_marks * right, negative_marks * wrong, " ", 0]
                student_ans = ["Student Ans", "Correct Ans", " ", "Student Ans", "Correct Ans"]
                sheet.append(no)
                sheet.append(marking)
                sheet.append(total_marks)
                sheet.append(blank_row)
                sheet.append(blank_row)
                sheet.append(student_ans)

                for i in range(0, 3):
                    new_row = [" ", answer_key[i], " ", " ", answer_key[25 + i]]
                    sheet.append(new_row)

                for i in range(3, 28):
                    new_row = [" ", answer_key[i]]
                    sheet.append(new_row)

                font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None, underline='none',
                            strike=False, color='000000')
                sheet['B8'].font = font
                sheet['C8'].font = font
                sheet['D8'].font = font
                sheet['E8'].font = font
                sheet['A9'].font = font
                sheet['A10'].font = font
                sheet['A11'].font = font
                font = Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                            strike=False, color='FF0000')
                sheet['C9'].font = font
                sheet['C10'].font = font
                sheet['C11'].font = font
                # green
                font = Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                            strike=False, color='00A300')
                sheet['B9'].font = font
                sheet['B10'].font = font
                sheet['B11'].font = font

                # blue
                font = Font(name='Calibri', size=11, bold=False, italic=False, vertAlign=None, underline='none',
                            strike=False, color='0000FF')
                sheet['E11'].font = font
                for i in range(15, 43):
                    s = "B" + str(i)
                    sheet[s].font = font
                    if i < 18:
                        s = "E" + str(i)
                        sheet[s].font = font

                font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None, underline='single',
                            strike=False, color='000000')
                sheet['C4'].font = font

                # black
                font = Font(name='Calibri', size=11, bold=True, italic=False, vertAlign=None, underline='none',
                            strike=False, color='000000')
                sheet['B5'].font = font
                sheet['B6'].font = font
                sheet['E5'].font = font
                sheet['A14'].font = font
                sheet['B14'].font = font
                sheet['D14'].font = font
                sheet['E14'].font = font

            wb.save(file_path)

    return redirect('Home')







