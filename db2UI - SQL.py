import time
import tkinter as tk
import easygui as e

import ibm_db_dbi
from xlwt import Workbook


class db2data():

    def __init__(self, cur_date, from_date):
        wb = Workbook()

        sheet1 = wb.add_sheet('Sheet 1')
        print(cur_date, from_date)
        if cur_date == '':
            curr_tm = time.localtime(time.time())
            curr_timestmp = str(curr_tm[0]) + '-' + str(curr_tm[1]) + '-' + str(curr_tm[2]) + '-' + str(curr_tm[3]) \
                            + '.' + str(curr_tm[4]) + '.' + str(curr_tm[5]) + '.000000'
        else:
            curr_timestmp = str(cur_date) + '-23.59.59.000000'

        if from_date == '':
            curr_tm = time.localtime(time.time())
            from_timestmp = str(curr_tm[0] - 1) + '-' + str(curr_tm[1]) + '-' + str(curr_tm[2]) + '-' \
                            + str(curr_tm[3]) + '.' + str(curr_tm[4]) + '.' + str(curr_tm[5]) + '.000000'
        else:
            from_timestmp = str(from_date) + '-00.00.00.000000'

        print(curr_timestmp, from_timestmp)
        sp = ' '

        conn = ibm_db_dbi.connect(dsn='DB2C', user='ln4', password='pridec18')

        # sql
        sql = "SELECT PPTILBUD_ID, CAST(TILBUDSDATO as char(10)), INTERR_ID, STAT, FILE_ID_HCP " \
              "FROM SCM474.PPTILBUD " \
              "WHERE INTERR_ID <> '" + sp + "' " \
              "AND   TMSTMP >=  '" + from_timestmp + "' " \
              "AND   TMSTMP <=  '" + curr_timestmp + "' " \
              "ORDER BY TILBUDSDATO,INTERR_ID,PPTILBUD_ID " \
              "WITH UR"
        #              "FETCH FIRST 2 ROWS ONLY " \

        print(sql)
        i = 0
        cur = conn.cursor()

        # EXECUTE ACTION QUERY BINDING PARAMS
        # stmt = cur.execute(sql, (sp, curr_timestmp, from_timestmp))
        stmt = cur.execute(sql)

        while True:
            rs = cur.fetchone()
            if not rs:
                break
            else:
                j = 0
                results = str(rs)
                results = results.replace('(', '', 1)
                results = results.replace("')", "',", 1)
                results = results + '\n'
                csv_rec = results.split(',')
                for field in csv_rec:
                    sheet1.write(i, j, field)
                    j += 1
                i += 1

        wb.save("D:\PyWorkSpace\pyprojects\Offerlist.xls")
        cur.close()
        conn.close()

class SimpleTable():
    entries = []

    def __init__(self, master):
        self.master = master
        mstr = tk.Frame(self.master)
        mstr.pack(side='top', fill='x', padx=10, pady=5)

        l = tk.Label(mstr, text='Start Date', width=30, anchor='w')
        l.grid(row=1, column=0, pady=5)
        e = tk.Entry(mstr)
        e.grid(row=1, column=1, pady=5)
        fmt = tk.Label(mstr, text='*CCYY-MM-DD', width=12, anchor='w')
        fmt.grid(row=1, column=2, pady=5)
        SimpleTable.entries.append(('Start Date', e))
        l1 = tk.Label(mstr, text='End Date', width=30, anchor='w')
        l1.grid(row=2, column=0, pady=5)
        e1 = tk.Entry(mstr)
        e1.grid(row=2, column=1, pady=5)
        fmt1 = tk.Label(mstr, text='*CCYY-MM-DD', width=12, anchor='w')
        fmt1.grid(row=2, column=2, pady=5)
        SimpleTable.entries.append(('End Date', e1))

        btn = tk.Button(mstr, text='Retrieve Data', command=self.popup, anchor='w')
        btn.grid(row=3, column=1, pady=5)

    def popup(self):
        entr = SimpleTable.entries
        remain_on_main = 0
        for data in entr:
            data_field = data[0]
            data_value = data[1].get()
            if data_field == 'Start Date':
                from_date = data_value
            else:
                cur_date = data_value

        if from_date > cur_date:
            e.msgbox('End date can not be greater then start date', 'Date Exception')
            remain_on_main += 1

        if remain_on_main == 0:
            db_data = db2data(cur_date, from_date)


if __name__ == "__main__":
    root = tk.Tk()
    m = SimpleTable(root)
    root.mainloop()