import csv
import xlsxwriter
import matplotlib.pyplot as plt
import io
import os

foldername = r"C:\Users\User\Desktop\second_degree\קורסים\סמסטר א\אלגוריתמים אבולציונים\final project\results_for_GA"
excel_folder_name = r"C:\Users\User\Desktop\second_degree\קורסים\סמסטר א\אלגוריתמים אבולציונים\final project\results_analysis"
for f in os.listdir(foldername):
    filename = foldername + "\\" + f
    xlsf = f[:-4] + ".xlsx"
    xlsfilename = excel_folder_name + "\\" + xlsf

    with(open(filename, "r")) as file_to_read:
        filereader = csv.reader(file_to_read)
        lines = list(filereader)

    cross_over_p = lines[0][1]
    cross_over_n = lines[1][1]
    mut_p = lines[2][1]
    avg_time = lines[3][1]
    total_time = lines[4][1]

    genID = []
    max_list = []
    avg_list = []
    median_list = []
    min_list = []
    hit_list = []


    def check_best_weights(weights):
        if weights[0] == max(weights):
            if weights[1] == max(weights[1:]):
                if weights[7] == max(weights[2:]):
                    temp_weights = [weights[i] for i in (2, 3, 4, 5, 6, 8, 9)]
                    if weights[2] == weights[3] == weights[4] == weights[5] == weights[6] == max(temp_weights):
                        if weights[8] > weights[9]:
                            return 1
        return 0


    for line in lines[6:]:
        genID.append(int(line[0]))
        max_list.append(float(line[1]))
        avg_list.append(float(line[2]))
        median_list.append(float(line[3]))
        min_list.append(float(line[4]))
        hit_list.append(check_best_weights(line[6:16]))

    plt.cla()
    plt.figure(0, figsize=(6.666, 5))
    plt.plot(max_list, label='max')
    plt.plot(avg_list, label='average')
    plt.plot(median_list, label='median')
    plt.plot(min_list, label='min')
    plt.title(f'cross-over p = {cross_over_p}, cross-over n = {cross_over_n}, mutation p = {mut_p}\navg time = {avg_time}, total time = {total_time}')
    plt.xticks(ticks=[genID[i] for i in range(0, len(genID), 10)])
    plt.xlabel('generation')
    plt.ylabel('fitness')
    leg = plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), fancybox=True, shadow=True, ncol=5)

    workbook = xlsxwriter.Workbook(xlsfilename)
    wks1 = workbook.add_worksheet("sheet1")
    for i, line in enumerate(lines):
        wks1.write_row(i, 0, line)

    imgdata = io.BytesIO()
    plt.savefig(imgdata, format="png", bbox_inches="tight")
    imgdata.seek(0)
    wks1.insert_image(0, 2, "", {'image_data': imgdata})

    plt.cla()
    plt.plot(hit_list, label='hits')
    plt.title(f'hits over generations')
    plt.xticks(ticks=[genID[i] for i in range(0, len(genID), 10)])
    plt.yticks(ticks=[0, 1])
    plt.xlabel('generation')
    plt.ylabel('hit')
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), fancybox=True, shadow=True, ncol=5)
    imgdata = io.BytesIO()
    plt.savefig(imgdata, format="png", bbox_inches="tight")
    imgdata.seek(0)
    wks1.insert_image(0, 12, "", {'image_data': imgdata})
    workbook.close()
