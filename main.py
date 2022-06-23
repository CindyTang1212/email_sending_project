import pandas as pd
import xlrd
from mailmerge import MailMerge
import datetime
from docx2pdf import convert
import os
import fitz
import shutil

path_original = os.getcwd()


def pdf_to_jpeg(input_dir):
    pdf_dir = []
    docunames = os.listdir(input_dir)
    os.chdir(input_dir)
    for docuname in docunames:
        if os.path.splitext(docuname)[1] == '.pdf':  # 目录下包含.pdf的文件
            pdf_dir.append(docuname)
    print(pdf_dir)
    for pdf in pdf_dir:
        doc = fitz.open(pdf)
        pdf_name = os.path.splitext(pdf)[0]
        page = doc[0]
        rotate = int(0)
        zoom_x = 1.5
        zoom_y = 1.5
        trans = fitz.Matrix(zoom_x, zoom_y).prerotate(rotate)
        pm = page.get_pixmap(matrix=trans, alpha=False)
        pm.save('%s.png' % pdf_name)
    os.chdir(path_original)


def get_basic_info(date):
    info = dict()
    employee_num = []
    employee_name = []
    signed_company = []
    entry_date = []
    transfer_date = []
    gangwei3 = []
    city = []
    for gonghao in df1[df1['试用期计划转正日期'] == date]['全球员工工号']:
        zhuangtai = df2[df2['员工工号'] == int(gonghao)]['流程状态'].item()
        if zhuangtai == 'COMPLETED':
            employee_num.append(int(gonghao))
            employee_name.append(df1[df1['全球员工工号'] == gonghao]['姓名'].item())
            signed_company.append(df1[df1['全球员工工号'] == gonghao]['合同签约单位'].item())
            entry_date.append(df1[df1['全球员工工号'] == gonghao]['入职日期'].item())
            transfer_date.append(df1[df1['全球员工工号'] == gonghao]['试用期计划转正日期'].item())
            gangwei3.append(df1[df1['全球员工工号'] == gonghao]['岗位族3（专业）'].item())
            city.append(df1[df1['全球员工工号'] == gonghao]['工作城市'].item())
    info['工号'] = employee_num
    info['姓名'] = employee_name
    info['合同主体'] = signed_company
    info['入职日期'] = entry_date
    info['计划转正日期'] = transfer_date
    info['岗位族3'] = gangwei3
    info['工作城市'] = city
    return info


def get_email_address(basic_info):
    emails = dict()
    emails['工号'] = basic_info['工号']
    emails['姓名'] = basic_info['姓名']
    personal_email = []
    manager_nums = []
    manager_names = []
    manager_email = []
    manager2_nums = []
    manager2_names = []
    manager2_email = []
    regional_names = []
    regional_email = []
    city_mentor_names = []
    city_mentor_email = []
    for num in basic_info['工号']:
        personal_email.append(df1[df1['全球员工工号'] == num]['公司邮箱'].item())
        manager_num = int(df1[df1['全球员工工号'] == num]['直属主管工号'].item())
        manager_nums.append(manager_num)
        manager_name = df1[df1['全球员工工号'] == num]['直属主管'].item()
        manager_names.append(manager_name)
        manager_email.append(df1[df1['全球员工工号'] == manager_num]['公司邮箱'].item())

        manager2_num = int(df1[df1['全球员工工号'] == num]['二级主管工号'].item())
        manager2_nums.append(manager2_num)
        manager2_name = df1[df1['全球员工工号'] == num]['二级主管'].item()
        manager2_names.append(manager2_name)
        manager2_email.append(df1[df1['全球员工工号'] == manager2_num]['公司邮箱'].item())

    emails['员工邮箱'] = personal_email

    emails['直属主管工号'] = manager_nums
    emails['直属主管'] = manager_names
    emails['直属主管邮箱'] = manager_email

    emails['二级主管工号'] = manager2_nums
    emails['二级主管'] = manager2_names
    emails['二级主管邮箱'] = manager2_email

    count = 0
    for gangwei3 in basic_info['岗位族3']:
        if gangwei3 == '临床业务':
            city = basic_info['工作城市'][count]
            print(city)
            regional_names.append(df3[df3['工作城市'] == city]['大区经理'].item())
            regional_email.append(df3[df3['工作城市'] == city]['大区经理邮箱'].item())
            city_mentor_names.append(df3[df3['工作城市'] == city]['带教组长'].item())
            if city != '上海':
                city_mentor_email.append(df3[df3['工作城市'] == city]['带教组长邮箱'].item())
            else:
                city_mentor_email.append('')
        else:
            regional_names.append('')
            regional_email.append('')
            city_mentor_names.append('')
            city_mentor_email.append('')
        count = count + 1
    emails['大区经理'] = regional_names
    emails['大区经理邮箱'] = regional_email
    emails['带教组长'] = city_mentor_names
    emails['带教组长邮箱'] = city_mentor_email
    emails = pd.DataFrame(emails)
    print(emails['带教组长邮箱'])
    emails['抄送'] = emails['直属主管邮箱'] + ',' + emails['二级主管邮箱'] + ',' + emails['大区经理邮箱'] + ',' + emails['带教组长邮箱']
    print(emails['抄送'])
    return emails


def get_emails_dict(emails):
    email_dict = dict()
    email_dict['employee_name'] = emails['employee_name']
    email_dict['personal_email'] = emails['personal_email']
    cc_email = []
    for key in emails.keys():
        if key in ['manager_email', 'manager2_email', 'regional_email', 'city_mentor']:
            for i in range(len(emails[key])):
                print(i, end=' : ')
                print(emails[key][i])


def drop_duplicates(emails):
    result = ''
    try:
        emails = list(set(emails.replace(';', ',').split(',')))
        for email in emails:
            if email != '':
                result += email + ';'
    except:
        pass
    return result[:-1]


def email_merge(docx_template, xlsx_file, output_path, sent_date):
    workbook = xlrd.open_workbook(xlsx_file)
    worksheet = workbook.sheet_by_index(0)
    nrow = worksheet.nrows
    for key in range(1, nrow):
        with MailMerge(docx_template) as doc:
            doc.merge(employee_name=str(worksheet.cell_value(key, 0)),
                      entry_date=str(worksheet.cell_value(key, 1)),
                      signed_company=str(worksheet.cell_value(key, 2)),
                      transfer_date=str(worksheet.cell_value(key, 3)),
                      sent_date=sent_date)
            num = int(worksheet.cell_value(key, 4))
            name = str(worksheet.cell_value(key, 0))
            output = output_path + '/{}-{}.docx'.format(num, name)
            doc.write(output)


def docx_to_pdf(file_path):
    for dirpath, dirnames, filenames in os.walk(file_path):
        for file in filenames:
            fullpath = os.path.join(dirpath, file)
            print('fullpath:' + fullpath)
            # convert("input.docx", "output.pdf")
            convert(fullpath, f"{fullpath}.pdf")  # 转换成pdf文件，但文件名是.docx.pdf，需要重新修改文件名

    # 修改文件名
    for dirpath, dirnames, filenames in os.walk(file_path):
        for fullpath in filenames:
            # print(fullpath)
            if '.pdf' in fullpath:
                fullpath_after = os.path.splitext(fullpath)[0]
                fullpath_after = os.path.splitext(fullpath_after)[0]
                fullpath_after = fullpath_after + '.pdf'
                fullpath_after = os.path.join(dirpath + '/' + fullpath_after)
                # print('fullpath_after:' + fullpath_after)
                fullpath = os.path.join(dirpath, fullpath)
                # print('fullpath:'+fullpath)
                os.rename(fullpath, fullpath_after)


def insert_into_excel(df, excel_path):
    df_original = pd.read_excel(excel_path)
    df_new = pd.concat([df_original, df], axis=0)
    df_new['入职日期'] = df_new['入职日期'].dt.date
    df_new['计划转正日期'] = df_new['计划转正日期'].dt.date
    df_new['实际转正日期'] = df_new['实际转正日期'].dt.date
    df_new['发送日期'] = df_new['发送日期'].dt.date
    df_new.to_excel(excel_path, index=False)


def deletefiles(path, keys, count=0):
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        print('正处理' + file_path)
        if os.path.isdir(file_path):
            for key in keys:
                if key in file:
                    try:
                        shutil.rmtree(file_path)
                        print('已删除' + file_path)
                        count += 1
                    except Exception as e:
                        print('未删除' + file_path)
                    continue
                else:
                    count = deletefiles(file_path, keys, count)
        elif os.path.isfile(file_path):
            for key in keys:
                if key in file:
                    try:
                        os.remove(file_path)
                        print('已删除' + file_path)
                        count += 1
                    except Exception as e:
                        print('未删除' + file_path)
                    continue
    return count


def format_date(date):
    return datetime.datetime.strptime(date, "%Y/%m/%d").strftime('%Y年%-m月%-d日')


if __name__ == '__main__':
    today_date = '2022/6/22'
    today_date_format = datetime.datetime.strptime(today_date, "%Y/%m/%d").strftime('%Y-%m-%d')
    sent_date = datetime.datetime.strptime(today_date, "%Y/%m/%d").strftime('%Y年%-m月%-d日')
    df1 = pd.read_csv('source_data/员工花名册-0622.csv')
    df2 = pd.read_csv('source_data/试用期转正报表包含流程中-0622.csv')
    df3 = pd.read_csv('source_data/大区经理及带教组长邮箱.csv')

    basic_info = get_basic_info(today_date)
    df_basic_info = pd.DataFrame(basic_info)
    df_basic_info['入职日期_发送版本'] = df_basic_info['入职日期'].apply(format_date)
    df_basic_info['计划转正日期_发送版本'] = df_basic_info['计划转正日期'].apply(format_date)
    result_path = 'result/' + today_date_format
    # if not os.path.exists(result_path):
    #     os.mkdir(result_path)
    # basic_info_output_path = result_path + '/basic_info_' + today_date_format + '.xlsx'
    # df_basic_info.to_excel(basic_info_output_path,
    #                        columns=['姓名', '入职日期_发送版本', '合同主体', '计划转正日期_发送版本', '工号'],
    #                        index=False)
    # docx_template_path = 'source_data/发送模版.docx'
    # pics_output_path = result_path + '/pics'
    # if not os.path.exists(pics_output_path):
    #     os.mkdir(pics_output_path)
    # email_merge(docx_template_path, basic_info_output_path, pics_output_path, sent_date)
    # docx_to_pdf(pics_output_path)
    # pdf_to_jpeg(pics_output_path)
    # folder_path = pics_output_path  # 要处理的文件夹
    # keys = ['pdf', 'docx']  # 要删除的关键字列表
    # count = deletefiles(folder_path, keys)
    # print('共删除{}项'.format(count))
    emails = get_email_address(basic_info)
    emails['抄送'] = emails['抄送'].apply(drop_duplicates)
    email_output_path = result_path + '/底表_' + today_date_format + '.xlsx'
    emails.to_excel(email_output_path,
                    columns=['工号', '姓名', '员工邮箱', '抄送'],
                    index=False)
    df_result = pd.concat([df_basic_info, emails], axis=1, join='inner')
    df_result = df_result.loc[:, ~df_result.columns.duplicated()]
    df_result['发送日期'] = today_date
    df_result['实际转正日期'] = df_result['计划转正日期']
    df_result['流程状态'] = 'COMPLETED'
    df_result = df_result[['工号',
                           '姓名',
                           '入职日期',
                           '计划转正日期',
                           '实际转正日期',
                           '流程状态',
                           '岗位族3',
                           '合同主体',
                           '直属主管工号',
                           '直属主管',
                           '直属主管邮箱',
                           '二级主管工号',
                           '二级主管',
                           '二级主管邮箱',
                           '工作城市',
                           '大区经理',
                           '大区经理邮箱',
                           '带教组长',
                           '带教组长邮箱',
                           '员工邮箱',
                           '抄送',
                           '发送日期']]
    df_result["入职日期"] = pd.to_datetime(df_result["入职日期"])
    df_result["计划转正日期"] = pd.to_datetime(df_result["计划转正日期"])
    df_result["实际转正日期"] = pd.to_datetime(df_result["实际转正日期"])
    df_result["发送日期"] = pd.to_datetime(df_result["发送日期"])
    insert_into_excel(df_result, 'result/发送记录.xlsx')
    df_sent_info = df_result[['工号', '姓名', '员工邮箱', '抄送']]
    df_sent_info['图片'] = df_sent_info['工号'].astype(str) + '-' + df_sent_info['姓名'] + '.png'
    df_sent_info.to_excel('result/' + today_date_format + '/发送信息.xlsx', columns=['工号', '姓名', '图片', '员工邮箱', '抄送'], index=False)
