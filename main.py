import json
import datetime as dt
import xlsxwriter as xlsxwriter


class colorstoprint:
    WAKE = '\033[95m'  # YELLOW
    REM = '\033[92m'  # GREEN
    DEEP = '\033[91m'  # RED
    LIGHT = '\033[94m'  # BLUE
    RESET = '\033[0m'  # RESET COLOR


with open('entireFile.json', 'r') as json_file:
    json_load = json.load(json_file)

# Create a workbook and add a worksheet.
sleepWorkBook = xlsxwriter.Workbook('SleepSheet.xlsx')
worksheet = sleepWorkBook.add_worksheet()

row = 0
col = 0
bold = sleepWorkBook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'border': 1})


for index, value in enumerate(json_load):
    totalMinOfDeep = 0
    totalMinOfLight = 0
    totalMinOfRem = 0
    totalMinOfWake = 0
    # convert string to date
    datetotest = json_load[index]['dateOfSleep']
    format = '%Y-%m-%d'
    datetoprint = dt.datetime.strptime(datetotest, format)

    inthoursOfSleep = int((json_load[index]['minutesAsleep']) / 60)  # find hours of sleep and convert to int
    hoursOfSleep = str(inthoursOfSleep)  # convert to string for printing
    remainder = str((json_load[index]['minutesAsleep']) % 60)  # find remainder

    # prints date of sleep data and length of the sleep
    print('\n\u001b[41;1m----------------------------------------\u001b[0m')
    print("  Date of Sleep:  " + datetoprint.strftime("%B %d, %Y"))
    print("  Length of Sleep:  " + hoursOfSleep + " Hours " + remainder + " Minutes")
    print('\u001b[41;1m----------------------------------------\u001b[0m\n')

    # start labeling rows and columns
    excelDate = datetoprint.strftime("%B %d, %Y")
    worksheet.write(row, col, excelDate)
    worksheet.write(row+2, col, "Light", bold)
    worksheet.write(row + 3, col, "Wake", bold)
    worksheet.write(row + 4, col, "REM", bold)
    worksheet.write(row + 5, col, "Deep", bold)

    fitbitData = json_load[index]['levels']['data']

    # loop to print the sleep stage data throughout the sleep cycle
    for x in fitbitData:
        timeOfSleepStage = x['dateTime']
        sleepStage = x['level']
        lengthOfStage = str((x['seconds']) / 60)

        dateofstage, justTime = timeOfSleepStage.rsplit('T', 1)

        formatExcelTime = sleepWorkBook.add_format({'num_format': 'hh:mm:ss AM/PM', 'bg_color': '#D9D9D9', 'bold': True, 'border': 1})
        formatWakeBG = sleepWorkBook.add_format({'bg_color': '#FFB77F'})
        formatLightBG = sleepWorkBook.add_format({'bg_color': '#99F3FF'})
        formatRemBG = sleepWorkBook.add_format({'bg_color': '#A2FF99'})
        formatDeepBG = sleepWorkBook.add_format({'bg_color': '#CC99FF'})

        format = ('%H:%M:%S.%f')
        timetoprint = dt.datetime.strptime(justTime, format)
        hour = timetoprint.hour

        worksheet.write(row+1, col + 1, justTime, formatExcelTime)
        col += 1

        print("------NEW STAGE------")
        if hour > 12:
            hour = hour - 12
            print("Time of Sleep Stage:  " + str(hour) + ":" + timetoprint.strftime("%M:%S%p"))

        elif hour == 0:
            hour = hour + 12
            print("Time of Sleep Stage:  " + str(hour) + ":" + timetoprint.strftime("%M:%S%p"))
        else:
            print("Time of Sleep Stage:  " + timetoprint.strftime("%H:%M:%S%p"))

        if sleepStage == 'wake':
            print("Sleep Stage:  " + colorstoprint.WAKE + sleepStage + colorstoprint.RESET)
            totalMinOfWake = totalMinOfWake + float(lengthOfStage)
            worksheet.write(row + 3, col, lengthOfStage + " min", formatWakeBG)
        elif sleepStage == 'light':
            print("Sleep Stage:  " + colorstoprint.LIGHT + sleepStage + colorstoprint.RESET)
            totalMinOfLight = totalMinOfLight + float(lengthOfStage)
            worksheet.write(row + 2, col, lengthOfStage + " min", formatLightBG)
        elif sleepStage == 'deep':
            print("Sleep Stage:  " + colorstoprint.DEEP + sleepStage + colorstoprint.RESET)
            totalMinOfDeep = totalMinOfDeep + float(lengthOfStage)
            worksheet.write(row + 5, col, lengthOfStage + " min", formatDeepBG)
        elif sleepStage == 'rem':
            print("Sleep Stage:  " + colorstoprint.REM + sleepStage + colorstoprint.RESET)
            totalMinOfRem = totalMinOfRem + float(lengthOfStage)
            worksheet.write(row + 4, col, lengthOfStage + " min", formatRemBG)
        print("Time in Stage:  " + lengthOfStage + " minutes")
    # prints the totals of each sleep stage
    print("\n")
    if totalMinOfLight >= 60:
        totalLightMin = int(totalMinOfLight/60)
        lightRemainder = totalMinOfLight%60
        print(colorstoprint.LIGHT + "Total Time of Light:  " + str(totalLightMin) + " hours " + str(lightRemainder) + " minutes" + colorstoprint.RESET)
    else:
        print(colorstoprint.LIGHT + "Total Time of Light:  " + str(totalMinOfLight) + " minutes" + colorstoprint.RESET)

    if totalMinOfWake >= 60:
        totalWakeMin = int(totalMinOfWake / 60)
        wakeRemainder = totalMinOfWake % 60
        print(colorstoprint.WAKE + "Total Time of Wake:  " + str(totalWakeMin) + " hours " + str(wakeRemainder) + " minutes" + colorstoprint.RESET)
    else:
        print(colorstoprint.WAKE + "Total Time of Wake:  " + str(totalMinOfWake) + " minutes" + colorstoprint.RESET)

    if totalMinOfRem >= 60:
        totalRemMin = int(totalMinOfRem / 60)
        remRemainder = totalMinOfRem % 60
        print(colorstoprint.REM + "Total Time of REM:  " + str(totalRemMin) + " hours " + str(remRemainder) + " minutes" + colorstoprint.RESET)
    else:
        print(colorstoprint.REM + "Total Time of REM:  " + str(totalMinOfRem) + " minutes" + colorstoprint.RESET)

    if totalMinOfDeep >= 60:
        totalDeepMin = int(totalMinOfDeep / 60)
        deepRemainder = totalMinOfDeep % 60
        print(colorstoprint.DEEP + "Total Time of Deep:  " + str(totalDeepMin) + " hours " + str(deepRemainder) + " minutes" + colorstoprint.RESET)
    else:
        print(colorstoprint.DEEP + "Total Time of Deep:  " + str(totalMinOfDeep) + " minutes" + colorstoprint.RESET)

    row = row + 7
    col = 0



sleepWorkBook.close()
