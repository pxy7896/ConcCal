# -*- coding:utf-8 -*-
import numpy
import os
import datetime
import xlwt
from pandas import read_excel
from pandas import read_csv
from shutil import copyfile
from shutil import rmtree
from cmath import sqrt

def getInput(inFile):
    result = [] 
    curDir = os.getcwd()
    i = datetime.datetime.now()
    workDir = "temp_" + str(i.hour) + str(i.minute)
    if workDir in os.listdir(os.getcwd()):
        rmtree(workDir)
    else:
        os.mkdir(workDir)
        os.chdir(workDir)
    inFile = inFile.lstrip().rstrip()
    # get all sheets
    data_xls = read_excel(inFile, sheet_name=None)
    name = os.path.splitext(inFile)[0].split("\\")[-1]
    standard_fold_list = []
    standard_conc_list = []
    sample_name_list = []
    sample_fold_list = []
    sample_conc_list = []
    sheet_name_list = []
    # separate files for each sheet
    for i in data_xls.keys():
        # for windows
        data_xls[i].to_csv(str(i) + ".csv", encoding='utf-8', index=False)
    for file in sorted(os.listdir(os.getcwd())):
        if os.path.splitext(file)[1] == '.csv':
            with open(file, 'r', encoding='utf-8') as file1:
                if file1.readline().find('Standards') == -1:
                    continue
                else:
                    sheet_name_list.append(file.split(".")[0])
                    standard_fold = []
                    standard_conc = []
                standard_flag = False
                sample_flag = False
                for line in file1:
                    #if line.find('Standards') != -1:    
                    if line.find("标准品") != -1:
                        standard_flag = True
                    elif line.find('Samples') != -1:
                        standard_flag = False
                        sample_fold = []
                        sample_conc = []
                    elif line.find("倍数") != -1:
                        sample_name = line.rstrip().split(",")[1:]
                        sample_flag = True
                    else:
                        if standard_flag:
                            info = line.rstrip().split(",")
                            standard_fold.append(float(info[0]))
                            standard_conc.append(round((float(info[1])+float(info[2])) / 2, 4))
                        if sample_flag:
                            info = line.rstrip().split(",")
                            sample_fold.append(int(info[0]))
                            i = 1
                            while i < len(info):
                                if info[i] == 'N':
                                    sample_conc.append(-1)
                                else:
                                    sample_conc.append(float(info[i]))
                                i += 1
                standard_fold_list.append(standard_fold)
                standard_conc_list.append(standard_conc)
                sample_name_list.append(sample_name)
                sample_fold_list.append(sample_fold)
                sample_conc_list.append(sample_conc)
    os.chdir(curDir)
    rmtree(workDir)
    return sheet_name_list, standard_fold_list, standard_conc_list, sample_name_list, sample_fold_list, sample_conc_list

# discard
def writeToExcel(outFile, sample_name, result_list, poly, error_list):
    workbook = xlwt.Workbook(encoding='ascii')
    worksheet = workbook.add_sheet('Result')
    # write samples
    for i in range(len(sample_name)):
        worksheet.write(i, 0, label=sample_name[i])
        worksheet.write(i, 1, label=result_list[i])
    # write expression
    parameters = poly['poly']
    expression = "Y = " + str(parameters[0]) + "*X*X + " + str(parameters[1]) + "*X + " + str(parameters[2])
    rSquare = "R-square: " + str(poly['r2'])
    i = len(sample_name) + 1
    worksheet.write(i, 0, label=expression)
    worksheet.write(i, 1, label=rSquare)
    # write error_list
    i += 2
    worksheet.write(i, 0, label='error_samples: "名称-稀释倍数-值"')
    i += 1
    for j in range(len(error_list)):
        info = error_list[j].split("-")
        rename = "-".join([sample_name[int(info[0])], info[1], info[2]])
        worksheet.write(i, 0, label=rename)
        i += 1
    workbook.save(outFile.split(".")[0] + '.xls')

# 3/4/2019 下午1修改版本的汇总格式
def writeToOutfile(idx, sample_name, result_list, poly, error_list, x_list):
    outFile = open("temp.csv", "a+")
    # write sheet name
    outFile.write(str(idx) + "\n")
    # write samples
    for i in range(len(sample_name)):
        outFile.write(",".join([sample_name[i], str(result_list[i])]) + "\n")
    parameters = poly['poly']
    expression = "Y = " + str(parameters[0]) + "*X*X + " + str(parameters[1]) + "*X + " + str(parameters[2])
    rSquare = "R-square: " + str(poly['r2'])
    outFile.write("\n" + ",".join([expression, rSquare]) + "\n")
    outFile.write("error_samples:name, fold, value\n")
    pre_name = "X"
    for i in range(len(error_list)):
        info = error_list[i].split("-")
        s_name = sample_name[int(info[0])]
        rename = ",".join([s_name, info[1], info[2]])
        if pre_name != s_name and pre_name != "X":
            outFile.write("\n")
        outFile.write(rename + "\n")
        pre_name = s_name
        
    outFile.write("\nX_samples:name, fold, value\n")
    pre_name = "X"
    for i in range(len(x_list)):
        info = x_list[i].split("-")
        s_name = sample_name[int(info[0])]
        rename = ",".join([s_name, info[1], info[2]])
        if pre_name != s_name and pre_name != "X":
            outFile.write("\n")
        outFile.write(rename + "\n")
        pre_name = s_name

    outFile.write("\n\n")
    outFile.close()

# 3/4/2019 下午2修改版本的汇总格式
def writeAsMatrix(idx, sample_name, sample_fold, result_list, poly, error_list, x_list):
    outFile = open("temp.csv", "a+")
    # write sheet name
    outFile.write(str(idx) + "\n")
    # write samples
    for i in range(len(sample_name)):
        outFile.write(",".join([sample_name[i], str(result_list[i])]) + "\n")
    parameters = poly['poly']
    expression = "Y = " + str(parameters[0]) + "*X*X + " + str(parameters[1]) + "*X + " + str(parameters[2])
    rSquare = "R-square: " + str(poly['r2'])
    outFile.write("\n" + ",".join([expression, rSquare]) + "\n\n")
    # write matrix
    sum_dict = {} 
    #print(error_list+x_list)
    for i in error_list + x_list:
        info = i.split("-")
        temp = sample_name[int(info[0])] + "-" + info[1]
        if temp not in sum_dict:
            sum_dict[temp] = info[2]

    outFile.write(",".join([" "] + sample_name) + "\n")
    for i in range(len(sample_fold)):
        fold = str(sample_fold[i])
        temp = []
        for m in sample_name:
            temp.append(sum_dict[m+"-"+fold])
        outFile.write(",".join([fold] + temp) + "\n")
    outFile.write("\n\n")
    outFile.close()



def polyfit(x, y, degree):
    results = {}
    coeffs = numpy.polyfit(x, y, degree)
    results['poly'] = coeffs.tolist()
    # r-squared
    p = numpy.poly1d(coeffs)
    # fit values, and mean
    yhat = p(x)                         
    ybar = numpy.sum(y)/len(y)          # average y
    ssreg = numpy.sum((yhat-ybar)**2)   # SSR
    sstot = numpy.sum((y - ybar)**2)    # SST
    results['r2'] = ssreg / sstot # R-square
    return results

def calConc(poly, fold, sample, x_min, x_max, y_min, y_max):
    parameters = poly['poly']
    fold = list(fold)
    sample = list(sample)
    len_sample = len(sample)
    len_fold = len(fold)
    step = len(sample) // len(fold)
    result_list = []
    error_sample = []
    x_list = []

    for i in range(step):
        result = []
        # for each sample, retrieve data
        for j in range(len_fold):
            num = sample[i + step * j]
            result.append(num)
        # get y
        cnt = 0
        sum = 0
        for j in range(len_fold):
            if result[j] < y_min or result[j] > y_max:
                #if result[j] != -1:
                error_sample.append(str(i)+"-"+str(fold[j])+"-N")
                continue
            else:
                cnt += 1
                # quadratic equation solvement
                x = -1
                a = parameters[0]
                b = parameters[1]
                c = parameters[2] - result[j]
                d = (b*b) - (4*a*c)
                sol1 = ((-b-sqrt(d))/(2*a)+0j).real
                sol2 = ((-b+sqrt(d))/(2*a)+0j).real
                #print(str(sol1))
                #print(str(sol2))
                if sol1 >= x_min and sol1 <= x_max:
                    x = sol1
                elif sol2 >= x_min and sol2 <= x_max:
                    x = sol2
                else:
                    print(sol1)
                    print(sol2)
                    error_sample.append(str(i)+"-"+str(fold[j])+"-N")
                    cnt -= 1
                    continue
                #print(x)
                temp_value = x * fold[j] / 1000
                sum += x * fold[j]
                x_list.append(str(i)+"-"+str(fold[j])+"-"+str(temp_value))
        if cnt <= 0: 
            ans = 0
        else:
            ans = sum / cnt / 1000
        result_list.append(round(ans, 3))
    return result_list, error_sample, x_list

def test(): 
    standard_fold, standard_conc, sample_name, sample_fold, sample_conc = getInput("D:\\Users\\pangxiaoyi\\Desktop\\样本数据o.xlsx")
    poly = polyfit(standard_fold, standard_conc, 2)
    result_list, error_list = calConc(poly, sample_fold, sample_conc, standard_conc[-2], standard_conc[0])
    writeToExcel('test2.xlsx', sample_name, result_list, poly, error_list)

if __name__ == '__main__':
    print("\nWelcome to ConcCalculator!\n")
    while input("Press q to quit or press c to continue\n") != "q":
        inputFile = input("Please input raw data file path: ")
        outputFile = input("\nOutput file name: ")
        inFile = inputFile.lstrip().rstrip()
        outFile = outputFile.lstrip().rstrip()
        sheet_name_list, standard_fold_list, standard_conc_list, sample_name_list, sample_fold_list, sample_conc_list = getInput(inFile)
        for i in range(len(sheet_name_list)):
            poly = polyfit(standard_fold_list[i], standard_conc_list[i], 2)
            result_list, error_list, x_list = calConc(poly, sample_fold_list[i], sample_conc_list[i], standard_fold_list[i][-1], standard_fold_list[i][0], standard_conc_list[i][-2], standard_conc_list[i][0])
            #writeToExcel(outFile, sample_name_list[i], result_list, poly, error_list)
            #writeToOutfile(sheet_name_list[i], sample_name_list[i], result_list, poly, error_list, x_list)
            writeAsMatrix(sheet_name_list[i], sample_name_list[i], sample_fold_list[i], result_list, poly, error_list, x_list)
        #csv = read_csv("temp.csv", encoding="ascii", engine='python')
        #csv.to_excel(outFile, sheet_name='Result')
        os.rename(os.path.join(os.getcwd(), 'temp.csv'),os.path.join(os.getcwd(), outFile.split(".")[0]+".csv"))
        #os.remove("temp.csv")
        print("\nDone\n")

    print("Bye!\n")

os.system("pause")
    