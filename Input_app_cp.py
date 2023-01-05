import pandas as pd
import PySimpleGUI as sg
import os
import xlwings as xl

import win32com.client as win32
from docx import Document
from openpyxl import load_workbook
import openpyxl

from datetime import date

sg.theme('DarkTeal9')
list_F=""
layout =[
        [sg.Text("Update infomation",background_color="red"),sg.Text("Link file Template"),sg.Input(key="link_word",size=(58,1)),sg.FileBrowse("link_word",file_types=[("file",".docx"),("file",".doc")],size=10,button_color="green")],
         [sg.Text("STT",size=(3,1)),sg.Spin([i for i in range (1,100)],key="STT",size=(4,1),background_color="grey"),sg.Text("Link File Excel",size=(19,1)),sg.Input(key="link_file",size=(28,1)),sg.FilesBrowse("Đường dẫn",file_types=[("file",".xlsx"),("file",".xls"),("file",".xlsm")],size=10,button_color="green"), sg.Text("1.Nhóm_dự_án",size=(18,1)),sg.Multiline(key="Nhóm_dự_án",size=(25,2)),sg.Text("2.Dự_án",size=(25,1)),sg.Multiline(key="Dự_án",size=(27,2)),sg.Text("3.Phòng_chủ_trì",size=(30,1)),sg.Combo(["Phòng KH-DT","Phòng Hạ Tầng","Phòng Truyền Dẫn","Tổ Xét Thầu","Phòng Kế Toán"],size=(28,4),key="Phòng_chủ_trì")],
        
        [sg.Text("4.PGĐ_phụ_trách",size=(30,1)),sg.Multiline(key="PGĐ_phụ_trách",size=(27,2)), sg.Text("5.Giấy_ủy_quyền",size=(30,1)),sg.Multiline(key="Giấy_ủy_quyền",size=(25,2)),sg.Text("6.Quyết_định_phê_duyệt_vốn",size=(25,1)),sg.Multiline(key="Quyết_định_phê_duyệt_vốn",size=(27,2)),sg.Text("7.Thời_gian_hợp_đồng",size=(30,1)),sg.Multiline(key="Thời_gian_hợp_đồng",size=(28,2))],
        
        [sg.Text("8.Giá_gói_thầu_Thiết_kế_trước_VAT",size=(30,1)),sg.Multiline(key="Giá_gói_thầu_Thiết_kế_trước_VAT",size=(27,2)), sg.Text("9.Giá_gói_thầu_Thiết_kế_sau_VAT",size=(30,1)),sg.Multiline(key="Giá_gói_thầu_Thiết_kế_sau_VAT",size=(25,2)),sg.Text("10.Giá_Thuế_gói_thầu_Thiết_kế_VAT",size=(25,1)),sg.Multiline(key="Giá_Thuế_gói_thầu_Thiết_kế_VAT",size=(27,2)), sg.Text("11.Giá_gói_thầu_Thiết_kế_bằng_chữ",size=(30,1)),sg.Multiline(key="Giá_gói_thầu_Thiết_kế_bằng_chữ",size=(28,2))],
        
        [sg.Text("12.Giá_gói_thầu_Thẩm_tra_trước_VAT",size=(30,1)),sg.Multiline(key="Giá_gói_thầu_Thẩm_tra_trước_VAT",size=(27,2)), sg.Text("13.Giá_gói_thầu_Thẩm_tra_sau_VAT",size=(30,1)),sg.Multiline(key="Giá_gói_thầu_Thẩm_tra_sau_VAT",size=(25,2)),sg.Text("14.Giá_thuế_gói_thầu_Thẩm_tra_VAT",size=(25,1)),sg.Multiline(key="Giá_thuế_gói_thầu_Thẩm_tra_VAT",size=(27,2)), sg.Text("15.Giá_gói_thầu_Thẩm_tra_bằng_chữ",size=(30,1)),sg.Multiline(key="Giá_gói_thầu_Thẩm_tra_bằng_chữ",size=(28,2))],
        
        [sg.Text("16.Thời_gian_bắt_đầu_LCNT_Thẩm_tra",size=(30,1)),sg.Multiline(key="Thời_gian_bắt_đầu_LCNT_Thẩm_tra",size=(27,2)), sg.Text("17.Thời_gian_thực_hiện_hợp_đồng_Thẩm_tra BCKTKT",size=(30,1)),sg.Multiline(key="Thời_gian_thực_hiện_hợp_đồng_Thẩm_tra BCKTKT",size=(25,2)),sg.Text("18.Thời_gian_bắt_đầu_LCNT_Thiết_kế",size=(25,1)),sg.Multiline(key="Thời_gian_bắt_đầu_LCNT_Thiết_kế",size=(27,2)), sg.Text("19.Thời_gian_thực_hiện_hợp_đồng_Thiết_kế_KS, lập BCKTKT",size=(30,1)),sg.Multiline(key="Thời_gian_thực_hiện_hợp_đồng_Thiết_kế_KS, lập BCKTKT",size=(28,2))],
         
        [sg.Text("20.Tờ_trình_phê_duyệt_TMĐT_kế_hoạch",size=(30,1)),sg.Multiline(key="Tờ_trình_phê_duyệt_TMĐT_kế_hoạch",size=(27,2)), sg.Text("21.Tờ_trình_phê_duyệt_dự_toán_và_KHLCNT_gói",size=(30,1)),sg.Multiline(key="Tờ_trình_phê_duyệt_dự_toán_và_KHLCNT_gói",size=(25,2)),sg.Text("22.Tờ_trình_phê_duyệt_BCKTKT__KHLCNT",size=(25,1)),sg.Multiline(key="Tờ_trình_phê_duyệt_BCKTKT__KHLCNT",size=(27,2)), sg.Text("23.Tờ_trình_phê_duyệt_HSMT",size=(30,1)),sg.Multiline(key="Tờ_trình_phê_duyệt_HSMT",size=(28,2))],
        
        [sg.Text("24.Ngày_trình_phê_duyệt_KQ_LCNT",size=(30,1)),sg.Multiline(key="Ngày_trình_phê_duyệt_KQ_LCNT",size=(27,2)), sg.Text("25.Tờ_trình_phê_duyệt_KQ_LCNT",size=(30,1)),sg.Multiline(key="Tờ_trình_phê_duyệt_KQ_LCNT",size=(25,2)),sg.Text("26.Báo_cáo_thẩm_định_dự_toán_và_KHLCNT_gói",size=(25,1)),sg.Multiline(key="Báo_cáo_thẩm_định_dự_toán_và_KHLCNT_gói",size=(27,2)), sg.Text("27.Báo_cáo_thẩm_định_BCKTKT",size=(30,1)),sg.Multiline(key="Báo_cáo_thẩm_định_BCKTKT",size=(28,2))],
        
        [sg.Text("28.Báo_cáo_thẩm_định_KQLCNT",size=(30,1)),sg.Multiline(key="Báo_cáo_thẩm_định_KQLCNT",size=(27,2)), sg.Text("29.Quyết_định_phê_duyệt_BCKTKT__KHLCNT",size=(30,1)),sg.Multiline(key="Quyết_định_phê_duyệt_BCKTKT__KHLCNT",size=(25,2)),sg.Text("30.Thời_gian_thực_hiện_dự_án_tư_vấn_trước",size=(25,1)),sg.Multiline(key="Thời_gian_thực_hiện_dự_án_tư_vấn_trước",size=(27,2)), sg.Text("31.Thời_gian_thực_hiện_trong_BCKTKT",size=(30,1)),sg.Multiline(key="Thời_gian_thực_hiện_trong_BCKTKT",size=(28,2))],
        
        [sg.Text("32.Quyết_định_phê_duyệt_thành_lâp_tổ_chuyên",size=(30,1)),sg.Multiline(key="Quyết_định_phê_duyệt_thành_lâp_tổ_chuyên",size=(27,2)), sg.Text("33.Số_KHLCNT",size=(30,1)),sg.Multiline(key="Số_KHLCNT",size=(25,2)),sg.Text("34.Số_TBMT",size=(25,1)),sg.Multiline(key="Số_TBMT",size=(27,2)), sg.Text("35.Ngày_đăng_TBMT",size=(30,1)),sg.Multiline(key="Ngày_đăng_TBMT",size=(28,2))],
        
        [sg.Text("36.Báo_cáo_thẩm_định_HSMT",size=(30,1)),sg.Multiline(key="Báo_cáo_thẩm_định_HSMT",size=(27,2)), sg.Text("37.Quyết_định_phê_duyệt_HSMT",size=(30,1)),sg.Multiline(key="Quyết_định_phê_duyệt_HSMT",size=(25,2)),sg.Text("38.Thời_gian_đóng_thầu",size=(25,1)),sg.Multiline(key="Thời_gian_đóng_thầu",size=(27,2)), sg.Text("39.Thời_gian_mở_thầu",size=(30,1)),sg.Multiline(key="Thời_gian_mở_thầu",size=(28,2))],
        
        [sg.Text("40.Ngày_đóng_thầu",size=(30,1)),sg.Multiline(key="Ngày_đóng_thầu",size=(27,2)), sg.Text("41.Ngày_mở_thầu",size=(30,1)),sg.Multiline(key="Ngày_mở_thầu",size=(25,2)),sg.Text("42.Thời_gian_chuẩn_bị_HSDT",size=(25,1)),sg.Multiline(key="Thời_gian_chuẩn_bị_HSDT",size=(27,2)), sg.Text("43.Báo_cáo_đánh_giá_HSDT",size=(30,1)),sg.Multiline(key="Báo_cáo_đánh_giá_HSDT",size=(28,2))],
        
        [sg.Text("44.Thời_gian_đánh_giá_HSDT",size=(30,1)),sg.Multiline(key="Thời_gian_đánh_giá_HSDT",size=(27,2)), sg.Text("45.Biên_bản_thương_thảo_hợp_đồng",size=(30,1)),sg.Multiline(key="Biên_bản_thương_thảo_hợp_đồng",size=(25,2)),sg.Text("46.Đơn_vị_lập_Thiết_kế",size=(25,1)),sg.Multiline(key="Đơn_vị_lập_Thiết_kế",size=(27,2)), sg.Text("47.Ngày báo_cáo_thẩm_tra_KTKT",size=(30,1)),sg.Multiline(key="Ngày báo_cáo_thẩm_tra_KTKT",size=(28,2))],
        
        [sg.Text("48.Số_Báo_cáo_thẩm_tra",size=(30,1)),sg.Multiline(key="Số_Báo_cáo_thẩm_tra",size=(27,2)), sg.Text("49.Đơn_vị_thẩm_tra_BCKTKT",size=(30,1)),sg.Multiline(key="Đơn_vị_thẩm_tra_BCKTKT",size=(25,2)),sg.Text("50.Tên_nhà_thầu_trúng_thầu",size=(25,1)),sg.Multiline(key="Tên_nhà_thầu_trúng_thầu",size=(27,2)), sg.Text("51.Địa_chỉ",size=(30,1)),sg.Multiline(key="Địa_chỉ",size=(28,2))],
        
        [sg.Text("52.Giá_trúng_thầu",size=(30,1)),sg.Multiline(key="Giá_trúng_thầu",size=(27,2)), sg.Text("53.Bằng_chữ_giá_trúng_thầu",size=(30,1)),sg.Multiline(key="Bằng_chữ_giá_trúng_thầu",size=(25,2)),sg.Text("54.Giá_gói_thầu_xây_dựng_phê_duyệt",size=(25,1)),sg.Multiline(key="Giá_gói_thầu_xây_dựng_phê_duyệt",size=(27,2)),sg.Text("55.Tên_gói_thầu_xây_dựng",size=(30,1)),sg.Multiline(key="Tên_gói_thầu_xây_dựng",size=(28,2))],
        
        [sg.Text("56.Tổng_mức_đầu_tư_kế_hoạch",size=(30,1)),sg.Multiline(key="Tổng_mức_đầu_tư_kế_hoạch",size=(27,2)), sg.Text("57.Tổng_mức_đầu_tư",size=(30,1)),sg.Multiline(key="Tổng_mức_đầu_tư",size=(25,2)),sg.Text("58.Bằng_chữ_TMĐT",size=(25,1)),sg.Multiline(key="Bằng_chữ_TMĐT",size=(27,2)), sg.Text("59.Mục_tiêu",size=(30,1)),sg.Multiline(key="Mục_tiêu",size=(28,2))],
        
        [sg.Text("60.Tiết_kiệm",size=(30,1)),sg.Multiline(key="Tiết_kiệm",size=(27,2)), sg.Text("61.Công văn làm rõ",size=(30,1)),sg.Multiline(key="Công văn làm rõ",size=(25,2)),sg.Text("62.Tên_gói_thầu_Thiết_kế",size=(25,1)),sg.Multiline(key="Tên_gói_thầu_Thiết_kế",size=(27,2)), sg.Text("63.Tên_gói_thầu_Thẩm_tra",size=(30,1)),sg.Multiline(key="Tên_gói_thầu_Thẩm_tra",size=(28,2))],
        
        [sg.Button("Create excel",size=15,button_color="green"),sg.Button("Save",size=10,button_color="green"),sg.Button("Motify",size=10,button_color="green"),sg.Button("Delete",size=10,button_color="green"),sg.Button("Reset",size=10,button_color="green"),sg.Button("Show Info",size=10,button_color="green"),sg.Button("Export show",size=10,button_color="green"),sg.Button("Export all",size=10,button_color="green"),sg.Button("Print Show",size=10,button_color="green"),sg.Button("Print all",size=10,button_color="green"),sg.Button("Exit",size=10,button_color="green")]
        
        ]


window=sg.Window("UPDATE DEVELOPER DUSA",layout)
def clear_input():
        for key in values:
                window[key](" ")
        return None

while True:
        event,values = window.read()
        #khi click vào Exit or dấu chéo để thoát
        if event == sg.WINDOW_CLOSED or event == "Exit":
                break
        elif event =="Create excel":
                wb =openpyxl.Workbook()
                ws = wb.active
                wb.save("Thông tin dự án.xlsx")
                df =pd.read_excel("Thông tin dự án.xlsx")
                sg.popup("Tạo thành công")
                sg.popup("Hãy tiến hành Save lại")
        #khi click vào Save        
        elif event == "Save":
                if values["link_file"] =="":
                        sg.popup("Hãy chọn file Excel")
                else:
                        a = values["link_file"]
                        df = pd.read_excel(a)
                        del values["link_file"]
                        del values["link_word"]
                        del values["link_word0"]
                        # del values["Đường dẫn"]
                        New_record = pd.DataFrame(values,index=[0])
                        df1=pd.concat([df,New_record],ignore_index=True)
                        df1.to_excel(a,index=False)
                        print(df1)
                        sg.popup("Đã lưu thành công")
                        clear_input()
        elif event =="Motify":
                if values["Dự_án"] =="" :
                        list_co=""
                        layout_1=[[sg.Text("Link File",size=(10,1)),sg.Input(key="Link",size=(25,1)),
                                  sg.FilesBrowse(file_types=[("file",".xlsx"),("file",".xls"),("file",".xlsm")])],
                                [sg.Text("STT Dự án",size=(3,1)),sg.InputCombo(list_co,key="STT Dự án",size=(4,1)),
                                 sg.Button("Nhập",size=10),sg.Cancel(size=10)]  
                        ]
                        window_1=sg.Window("Xuất Thông Tin Dự Án (DUSA)",layout_1)
                        
                        while True:
                                event_1,values_1 = window_1.read()
                                if event_1 in (sg.WIN_CLOSED,"Cancel"):
                                        window_1.close()
                                        break
                                elif event_1 =="Nhập":
                                        if values_1["STT Dự án"] =="" and values_1["Link"] =="" :
                                                sg.popup("Hãy nhập Link và tên dự án")
                                        elif values_1["STT Dự án"] ==""  :
                                                sg.popup("Hãy nhập STT dự án")
                                                a = values_1["Link"]
                                                df = pd.read_excel(a)
                                                df_stt = df["STT"].to_list()
                                                window_1["STT Dự án"].Update(values=df_stt)
                                        elif values_1["Link"] =="" :
                                                sg.popup("Hãy nhập Link dự án")
                                        else:
                                                file_name = values_1["STT Dự án"]
                                                a = values_1["Link"]
                                                df = pd.read_excel(a)
                                                indexa = df.loc[df["STT"]==int(file_name)].index.to_list()[0]
                                                df_name = df.iloc[indexa]
                                                df_dict=df_name.to_dict()
                                                window_1.close()
                                                for key,value in df_dict.items():
                                                        window[key].Update(value)
                                               
                                                        
                else:
                        if values["link_file"] =="" :
                                sg.popup("Hãy chọn link file")
                        else:
                                file_name = values["STT"]
                                print(file_name)
                                a = values["link_file"]
                                df = pd.read_excel(a)
                                print(df["STT"])
                                indexa = df.loc[df["STT"]==int(file_name)].index.to_list()[0]
                                head_list=list(df.columns.values)
                                print(head_list)
                                for key in head_list:
                                        df.loc[indexa,key]=values[key]
                                df.to_excel(a,index=False)
                                sg.popup("Chỉnh sửa thành công")
                                clear_input()
                        
        #khi click vào reset
        elif event == "Reset":
                clear_input()
        
        #khi click vào Delete
        elif event =="Delete":
                if values["Dự_án"] =="" :
                        list_co=""
                        layout_1=[[sg.Text("Link File",size=(10,1)),sg.Input(key="Link",size=(25,1)),
                                  sg.FilesBrowse(file_types=[("file",".xlsx"),("file",".xls"),("file",".xlsm")])],
                                [sg.Text("STT Dự án",size=(3,1)),sg.InputCombo(list_co,key="STT Dự án",size=(4,1)),
                                 sg.Button("Nhập",size=10),sg.Cancel(size=10)]  
                        ]
                        window_1=sg.Window("Xuất Thông Tin Dự Án (DUSA)",layout_1)
                        
                        while True:
                                event_1,values_1 = window_1.read()
                                if event_1 in (sg.WIN_CLOSED,"Cancel"):
                                        window_1.close()
                                        break
                                elif event_1 =="Nhập":
                                        if values_1["STT Dự án"] =="" and values_1["Link"] =="" :
                                                sg.popup("Hãy nhập Link và tên dự án")
                                        elif values_1["STT Dự án"] ==""  :
                                                sg.popup("Hãy nhập STT dự án")
                                                a = values_1["Link"]
                                                df = pd.read_excel(a)
                                                df_stt = df["STT"].to_list()
                                                window_1["STT Dự án"].Update(values=df_stt)
                                        elif values_1["Link"] =="" :
                                                sg.popup("Hãy nhập Link dự án")
                                        else:
                                                file_name = values_1["STT Dự án"]
                                                a = values_1["Link"]
                                                df = pd.read_excel(a)
                                                indexa = df.loc[df["STT"]==int(file_name)].index.to_list()[0]
                                                df_name = df.iloc[indexa]
                                                df_dict = df_name.to_dict()
                                                window_1.close()
                                                for key,value in df_dict.items():
                                                        window[key].Update(value)
                                                
                                                        
                else:
                        if values["link_file"] =="" :
                                sg.popup("Hãy chọn link file")
                        else:
                                file_name = values["STT"]
                                print(file_name)
                                a = values["link_file"]
                                print(a)
                                df = pd.read_excel(a)
                                print(df["STT"])
                                indexa = df.loc[df["STT"]==int(file_name)].index.to_list()[0]
                                delete_df = df.drop(indexa)
                                delete_df.to_excel(a,index=False)
                                sg.popup("Xóa thành công")
                                clear_input()
                                
        elif event =="Show Info":
                if values["Dự_án"] =="" :
                        sg.popup("Hãy nhập số liệu để show")
                        list_co=""
                        layout_1=[[sg.Text("Link File",size=(10,1)),sg.Input(key="Link",size=(25,1)),
                                  sg.FilesBrowse(file_types=[("file",".xlsx"),("file",".xls"),("file",".xlsm")])],
                                [sg.Text("STT Dự án",size=(3,1)),sg.InputCombo(list_co,key="STT Dự án",size=(4,1)),
                                 sg.Button("Nhập",size=10),sg.Cancel(size=10)]  
                        ]
                        window_1=sg.Window("Xuất Thông Tin Dự Án (DUSA)",layout_1)
                        
                        while True:
                                event_1,values_1 = window_1.read()
                                if event_1 in (sg.WIN_CLOSED,"Cancel"):
                                        window_1.close()
                                        break
                                elif event_1 =="Nhập":
                                        if values_1["STT Dự án"] =="" and values_1["Link"] =="" :
                                                sg.popup("Hãy nhập Link và tên dự án")
                                                
                                        elif values_1["STT Dự án"] ==""  :
                                                sg.popup("Hãy nhập STT dự án")
                                                a = values_1["Link"]
                                                df = pd.read_excel(a)
                                                df_stt = df["STT"].to_list()
                                                window_1["STT Dự án"].Update(values=df_stt)
                                                
                                                
                                        elif values_1["Link"] =="" :
                                                sg.popup("Hãy nhập Link dự án")
                                                
                                        else:
                                                file_name = values_1["STT Dự án"]
                                                a = values_1["Link"]
                                                df = pd.read_excel(a)
                                                indexa = df.loc[df["STT"]==int(file_name)].index.to_list()[0]
                                                df_name = df.iloc[indexa]
                                                df_dict = df_name.to_dict()
                                                window_1.close()
                                                for key,value in df_dict.items():
                                                        window[key].Update(value)
                                                
                                                
                                                                
                else:
                        if values["link_file"] =="" :
                                sg.popup("Hãy nhập link để thao tác")
                                
        elif event =="Export all" or event == "Export show" or event =="Print Show" or event =="Print all":
                if values["Dự_án"]=="" and values["link_word"]=="":
                        sg.popup("Hãy chọn link vào")
                elif values["Dự_án"]=="":
                        sg.popup("Hãy Show Info")
                elif values["link_word"]=="":
                        sg.popup("Hãy nhập link file word")
                else:
                        if event =="Export all" :
                                
                                direct_disks=os.getcwd()
                                word_directory = values["link_word"]
                                excel_directory = values["link_file"]
                                
                                """
                                Create a Word application instance
                                """
                                wordApp = win32.Dispatch('Word.Application')
                                wordApp.Visible = True

                                """
                                Open Word Template + Open Data Source
                                """
                                sourceDoc = wordApp.Documents.Open(word_directory)
                                mail_merge = sourceDoc.MailMerge
                                mail_merge.OpenDataSource(excel_directory)

                                record_count = mail_merge.DataSource.RecordCount

                                """
                                Perform Mail Merge
                                """
                                for i in range(1, record_count + 1):
                                        mail_merge.DataSource.ActiveRecord = i
                                        mail_merge.DataSource.FirstRecord = i
                                        mail_merge.DataSource.LastRecord = i

                                        mail_merge.Destination = 0
                                        mail_merge.Execute(False)
                
                                        continue
                        elif event =="Print all":
                                direct_disks=os.getcwd()
                                word_directory = values["link_word"]
                                excel_directory = values["link_file"]
                                
                                """
                                Create a Word application instance
                                """
                                wordApp = win32.Dispatch('Word.Application')
                                wordApp.Visible = True

                                """
                                Open Word Template + Open Data Source
                                """
                                sourceDoc = wordApp.Documents.Open(word_directory)
                                mail_merge = sourceDoc.MailMerge
                                mail_merge.OpenDataSource(excel_directory)

                                record_count = mail_merge.DataSource.RecordCount

                                """
                                Perform Mail Merge
                                """
                                for i in range(1, record_count + 1):
                                        mail_merge.DataSource.ActiveRecord = i
                                        mail_merge.DataSource.FirstRecord = i
                                        mail_merge.DataSource.LastRecord = i

                                        mail_merge.Destination = 1
                                        mail_merge.Execute(False)
                
                                        continue
                                

                                # # get record value
                                # base_name = mail_merge.DataSource.DataFields('Name of Recipient'.replace('', '_')).Value()

                                # targetDoc = wordApp.ActiveDocument

                                # """
                                # Save Files in Word Doc and PDF
                                # """
                                # targetDoc.SaveAs2(os.path.join(direct_disks, base_name + '.docx'), 16)
                                # targetDoc.ExportAsFixedFormat(os.path.join(direct_disks, base_name), exportformat=17)
                                
                                # """
                                # Close target file
                                # """
                                # targetDoc.Close(False)
                                # targetDoc = None
                                
                                # sourceDoc.MailMerge.MainDocumentType = -1
                        elif event == "Export show":
                                direct_disks=os.getcwd()
                                word_directory = values["link_word"]
                                excel_directory = values["link_file"]
                                
                                """
                                Create a Word application instance
                                """
                                wordApp = win32.Dispatch('Word.Application')
                                wordApp.Visible = True

                                """
                                Open Word Template + Open Data Source
                                """
                                sourceDoc = wordApp.Documents.Open(word_directory)
                                mail_merge = sourceDoc.MailMerge
                                mail_merge.OpenDataSource(excel_directory)

                                record_count = mail_merge.DataSource.RecordCount

                                """
                                Perform Mail Merge
                                """
                                file_name = values["STT"]
                                
                                mail_merge.DataSource.ActiveRecord = int(file_name)
                                mail_merge.DataSource.FirstRecord = int(file_name)
                                mail_merge.DataSource.LastRecord = int(file_name)
                                
                                mail_merge.Destination = 0
                                mail_merge.Execute(True)
                                
                        elif event =="Print Show" :
                                direct_disks=os.getcwd()
                                word_directory = values["link_word"]
                                excel_directory = values["link_file"]
                                
                                """
                                Create a Word application instance
                                """
                                wordApp = win32.Dispatch('Word.Application')
                                wordApp.Visible = True

                                """
                                Open Word Template + Open Data Source
                                """
                                sourceDoc = wordApp.Documents.Open(word_directory)
                                mail_merge = sourceDoc.MailMerge
                                mail_merge.OpenDataSource(excel_directory)

                                record_count = mail_merge.DataSource.RecordCount

                                """
                                Perform Mail Merge
                                """
                                file_name = values["STT"]
                                
                                mail_merge.DataSource.ActiveRecord = int(file_name)
                                mail_merge.DataSource.FirstRecord = int(file_name)
                                mail_merge.DataSource.LastRecord = int(file_name)
                                
                                mail_merge.Destination = 1
                                mail_merge.Execute(False)
                                
                        
                                        
                                
                                
                                # # get record value
                                # base_name = mail_merge.DataSource.DataFields('Name of Recipient').Value

                                # targetDoc = wordApp.ActiveDocument

                                # """
                                # Save Files in Word Doc and PDF
                                # """
                                # targetDoc.SaveAs2(os.path.join(direct_disks, base_name + '.docx'), 16)
                                # targetDoc.ExportAsFixedFormat(os.path.join(direct_disks, base_name), exportformat=17)
                                
                                # """
                                # Close target file
                                # """
                                # targetDoc.Close(False)
                                # targetDoc = None
                                
                                # sourceDoc.MailMerge.MainDocumentType = -1
                                
                                        
        
               
                        
                        
                        
                
                                        
                                                
                                                
                                        
                                        
                                
                        
                                
                        
                       
                        
                        
                
        