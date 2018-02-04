from bs4 import BeautifulSoup
import requests
import xlsxwriter

url = 'http://results.vtu.ac.in/cbcs_17/result_page.php'


def getResult(usn):
    payload = {'usn': usn}
    response = requests.post(url, data=payload)

    soup = BeautifulSoup(response.text, 'html.parser')
    div_blocks = soup.findAll('div', attrs={'class': 'col-md-12'})
    name = div_blocks[3].findAll('td')[3].text[2:]
    marks = []
    for i in range(1, 9):
        marks.append(div_blocks[4].findAll('tr')[i].findAll('td')[4].text)

    sgpa = '{0:.2f}'.format(calcSGPA(marks))

    return [usn, name, sgpa]


def writeToXlsx(data, branch, num):
    workbook = xlsxwriter.Workbook('Results.xlsx')
    worksheet = []

    wbformat = workbook.add_format()
    wbformat.set_font_name('Ubuntu Mono')

    for i in range(len(branch)):
        worksheet.append(workbook.add_worksheet(branch[i]))

        worksheet[i].set_column(0, 0, 10)
        worksheet[i].set_column(1, 1, 35)
        worksheet[i].set_column(2, 2, 4)

        worksheet[i].write(0, 0, "USN", wbformat)
        worksheet[i].write(0, 1, "NAME", wbformat)
        worksheet[i].write(0, 2, "SGPA", wbformat)

        offset = 0
        for n in range(i):
            offset += num[i]

        for j in range(offset, num[i]):
            for k in range(3):
                worksheet[i].write(j + 1, k, data[j][k], wbformat)

    workbook.close()


def getGrade(mark):
    if mark in range(90, 101):
        return 10
    elif mark in range(80, 90):
        return 9
    elif mark in range(70, 80):
        return 8
    elif mark in range(60, 70):
        return 7
    elif mark in range(50, 60):
        return 6
    elif mark in range(45, 50):
        return 5
    elif mark in range(40, 45):
        return 4
    else:
        return 0


def calcSGPA(marks):
    sgpa = 0

    for i in range(6):
        sgpa = sgpa + getGrade(int(marks[i])) * 4
    for i in range(6, 8):
        sgpa = sgpa + getGrade(int(marks[i])) * 2

    return sgpa / 28


def main():
    branches = ['CS', 'IS', 'EC', 'ME']
    numStuds = [195, 135, 195, 195]
    result = []
    retry = []

    for i in range(4):
        print(branches[i])
        for j in range(1, numStuds[i]):
            usn = ('1PE15' + branches[i] + '{0:03d}').format(j)

            try:
                result.append(getResult(usn))
                print(j)
            except IndexError:
                pass
            except:
                retry.append([i, j, usn])
                print(str(j) + " E")
                pass

    for r in retry:
        print("\nRetrying...")

        offset = 0

        for i in range(r[0]):
            offset += numStuds[i]
        offset += r[1]

        try:
            result.insert(offset, getResult(r[2]))
            print(r[2])
        except Exception as e:
            print(e)
            pass

    writeToXlsx(result, branches, numStuds)


if __name__ == '__main__':
    main()
