import io
from django.http import Http404
from django.shortcuts import render, get_object_or_404
from django.http import HttpResponse, HttpResponseRedirect, FileResponse
from django.http import JsonResponse
from django.template import loader
from django.urls import reverse
from django.views import generic
from django.utils import timezone

from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import pandas as pd
import win32com.client  #Module 설치: pywin32
# jinja2 설치 필요
import xlwt
import xlsxwriter

from JobReport.model_SubmissionStatus import ModelSubmissionStatus
from JobReport.model_PredictPeriod import ModelPredictPeriod
from JobReport.model_WorldMap import *
from JobReport.models import *

from JobReport.model_JobReportMan import DC_JobReportMan
from JobReport.model_JobReportEqu import DC_JobReportEqu
from JobReport.model_JobReportQnt import DC_JobReportQnt
from JobReport.model_JobReportPlnt import DC_JobReportPlnt


def view_world_map(request):
    df = chart_data_for_world_map()
    chart = chart_world_map(df)
    context = {
        "title": "■ 출역 현황 World Map (누계)",
        "chart": chart,
    }
    return render(request, 'JobReport/JobReportWorldMap.html', context)


def view_daily_report_summary(request):
    domestic = ModelSubmissionStatus("DOMESTIC")
    df1 = domestic.resultBonsaExcel_Create()
    oversea = ModelSubmissionStatus("OVERSEA")
    df2 = oversea.resultBonsaExcel_Create()
    df = pd.concat([df1, df2])
    df = df.reset_index()
    df_all = df

    if request.method == 'GET':

        df = df[['현장명','업체','1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']]
        df = df.set_index(['현장명', '업체'])

        df = pd.concat(
            [d.append(d.sum().rename((k, 'Total'))) for k, d in df.groupby(level=0)]).append(
            df.sum().rename(('Grand', 'Total')))

        is_exist = True
        site_code = ""
        dom_code = "전체"
        bonbu_code = "전체"

    elif request.method == 'POST':

        site_code = request.POST.get('site_code', "")
        site_code = site_code.upper()
        print("현장코드 : ", site_code)
        dom_code = request.POST.get('dom_code', "전체")
        print("국내/해외 : ", dom_code)
        bonbu_code = request.POST.get('bonbu_code', "전체")
        print("본부 : ", bonbu_code)

        # 해당 현장이 있으면, 해당 현장의 데이터

        if bonbu_code == "토목사업본부":
            df3 = df[df['본부'].str.contains("토목사업본부", regex=False)]
            df4 = df[df['본부'].str.contains("Infrastructure Division", regex=False)]
            df = pd.concat([df3, df4])
        elif bonbu_code != "전체":
            df = df[df['본부'].str.contains(bonbu_code, regex=False)]
        if dom_code != "전체":
            df = df[df['국내해외'] == dom_code]
        if site_code != "":
            df = df[df['현장명'].str.contains(site_code, regex=False)]

        # if dataframe is empty, show all data
        if len(df.index) > 0:
            is_exist = True
        else:
            is_exist = False
            df = df_all
            site_code = ""
            dom_code = "전체"
            bonbu_code = "전체"


        # 본부, 국내해외 제외하고 데이터를 보여줌
        df = df[['현장명','업체','1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']]
        df = df.set_index(['현장명', '업체'])

        df = pd.concat(
            [d.append(d.sum().rename((k, 'Total'))) for k, d in df.groupby(level=0)]).append(
            df.sum().rename(('Grand', 'Total')))

    # format numeric data
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    for col_name in numeric_columns:
        df[col_name] = df[col_name].apply(lambda x: "{:,.0f}".format(x))
    df = df.replace("0", "-")

    json_records = df.reset_index().to_json(orient='records')
    data_all = json.loads(json_records)

    context = {
        "title": "■ 일일 제출 현황",
        "date": datetime.today(),
        "data": data_all,
        "site_code": site_code,
        "bonbu_code": bonbu_code,
        "dom_code": dom_code,
        "is_exist": is_exist,
    }
    return render(request, 'JobReport/JobReportSummary.html', context)


def view_site_timezone(request):
    if request.method == 'POST':
        site_code = request.POST.get('site_code', False)
        site_code = site_code.upper()
        if JobMstSite.objects.filter(현장코드__iexact = site_code):
            item = JobMstSite.objects.filter(현장코드__iexact = site_code).values()
            is_exist = True
        else:
            is_exist = False
            item = JobMstSite.objects.all().values()
    else:
        is_exist = True
        site_code = "all"
        item = JobMstSite.objects.all().values()
    df = pd.DataFrame(item)
    # function
    df_html = arrange_dataframe_to_html_without_index(df)
    # text alignment / column size :
    alignment_columns = """
        <style>
            /* 각 컬럼별 너비 및 정렬 조정 옵션 */

        .col0 {width: 200px; text-align: center;}
        .col1 {width: 400px; text-align: center;}
        .col2 {width: 200px; text-align: center;}
        .col3 {width: 200px; text-align: center;}
        .col4 {width: 300px; text-align: center;}

        </style>
    """
    context = {
        "title": "■ 현장 별칭 및 타임존",
        "date": datetime.today(),
        "alignment_columns": alignment_columns,
        "df": df_html,
        "site_code": site_code,
        "is_exist": is_exist,
    }
    return render(request, 'JobReport/JobReportMainPD.html', context)


def view_site_summary(request):
    if request.method == 'POST':
        site_code = request.POST.get('site_code', False)
        site_code = site_code.upper()
        if PrdMstSite.objects.filter(현장코드__iexact = site_code):
            item = PrdMstSite.objects.filter(현장코드__iexact = site_code).values()
            is_exist = True
        else:
            is_exist = False
            item = PrdMstSite.objects.all().values()
    else:
        is_exist = True
        site_code = "all"
        item = PrdMstSite.objects.all().values()
    df = pd.DataFrame(item)
    # function
    df_html = arrange_dataframe_to_html_without_index(df)
    # text alignment / column size :
    alignment_columns = """
        <style> 
            /* 각 컬럼별 너비 및 정렬 조정 옵션 */

        .col0 {width: 100px; text-align: center;}
        .col1 {width: 150px; text-align: center;}
        .col2 {width: 100px; text-align: center;}
        .col3 {width: 300px; text-align: center;}
        .col4 {width: 100px; text-align: center;}
        .col5 {width: 100px; text-align: center;}
        .col6 {width: 100px; text-align: center;}
        .col7 {width: 100px; text-align: center;}
        .col8 {width: 100px; text-align: center;}
        .col9 {width: 100px; text-align: center;}
        .col10 {width: 100px; text-align: center;}
        .col11 {width: 100px; text-align: center;}
        .col12 {width: 100px; text-align: center;}

        </style>
    """
    context = {
        "title": "■ 현장 현황",
        "date": datetime.today(),
        "alignment_columns": alignment_columns,
        "df": df_html,
    }
    return render(request, 'JobReport/JobReportMainPD.html', context)


def arrange_dataframe_to_html_with_index(df):
    # function
    s = arrange_dataframe(df)
    # render a styled df to html
    df_html = s.render()
    return df_html

def arrange_dataframe_to_html_without_index(df):
    # function
    s = arrange_dataframe(df)
    # hide index for delete th in front of each row
    df_html = s.hide_index().render()
    return df_html

def arrange_dataframe(df):
    # add style for html
    css = pd.DataFrame([["other-class"]])
    # set class name
    s = df.style.set_td_classes(css)
    # round number : 1.233 -> 1
    # s = s.set_precision(0)
    s = s.format(precision=0)
    # set data frame class with "dataframe"
    s = s.set_table_attributes('class="dataframe"')
    # # render a styled df to html
    # df_html = s.render()
    # # hide index for delete th in front of each row
    # df_html = s.hide_index().render()

    # # to_html : option -> easy to work but couldn't set class attribute
    # df_html = df.reset_index().to_html(col_space=120, justify='inherit', border=0, index=False)
    # df_html = df_html.replace(before, after)

    return s


def view_excel_download(request):
    # operation data
    domestic = ModelSubmissionStatus("DOMESTIC")
    df1 = domestic.resultBonsaExcel_Create()
    oversea = ModelSubmissionStatus("OVERSEA")
    df2 = oversea.resultBonsaExcel_Create()
    df = pd.concat([df1, df2])
    df = df.reset_index()

    # save in memory by pandas / io
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer) as writer:
        df.to_excel(writer)
    buffer.seek(0)

    # response
    response = HttpResponse(buffer.read(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = "attachment; filename=dailyJobReportSummary.xlsx"

    # memory close
    buffer.close()

    return response


def view_summary_ilbo_by_site(request):
    # operation data
    domestic = ModelSubmissionStatus("DOMESTIC")
    df1 = domestic.resultBonsaExcel_Create()
    oversea = ModelSubmissionStatus("OVERSEA")
    df2 = oversea.resultBonsaExcel_Create()
    df = pd.concat([df1, df2])
    df = df.reset_index()
    df_all = df.copy()

    ### 본부별
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                       '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    df_bonbu = df_all.groupby(['본부'])[numeric_columns].sum()
    df = df_bonbu.copy()
    # format numeric data
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                       '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    for col_name in numeric_columns:
        df[col_name] = df[col_name].apply(lambda x: "{:,.0f}".format(x))
    df = df.replace("0", "-")
    # JSON setting
    json_records = df.reset_index().to_json(orient='records')
    data_all = json.loads(json_records)


    ### 국내해외별
    df_dom = df_all.groupby(['국내해외'])[numeric_columns].sum()
    df = df_dom.copy()
    # format numeric data
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                       '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    for col_name in numeric_columns:
        df[col_name] = df[col_name].apply(lambda x: "{:,.0f}".format(x))
    df = df.replace("0", "-")
    # JSON setting
    json_records = df.reset_index().to_json(orient='records')
    data_all_3 = json.loads(json_records)


    ### 현장별
    df_site = df_all.groupby(['현장명'])[numeric_columns].sum()
    df = df_site.copy()
    # format numeric data
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                       '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    for col_name in numeric_columns:
        df[col_name] = df[col_name].apply(lambda x: "{:,.0f}".format(x))
    df = df.replace("0", "-")
    # JSON setting
    json_records = df.reset_index().to_json(orient='records')
    data_all_2 = json.loads(json_records)



    ## 차트 그리기

    import plotly.express as px
    df_bonbu = df_bonbu.rename({'5_인원(누계)': '누계인원'}, axis=1)
    # df_bonbu['누계인원'] = df_bonbu['누계인원'].apply(lambda x: float(x))
    df_bonbu = df_bonbu.reset_index()
    fig = px.pie(df_bonbu, values='누계인원', names='본부', color_discrete_sequence=px.colors.sequential.RdBu)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    chart_html = fig.to_html(include_plotlyjs='cdn')
    chart = re.findall('<body>(.*?)</body>', chart_html, re.DOTALL)
    chart = "".join(chart)
    chart_1 = chart

    df_dom = df_dom.rename({'5_인원(누계)': '누계인원'}, axis=1)
    df_dom = df_dom.reset_index()
    fig = px.pie(df_dom, values='누계인원', names='국내해외', color_discrete_sequence=px.colors.sequential.RdBu)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    chart_html = fig.to_html(include_plotlyjs='cdn')
    chart = re.findall('<body>(.*?)</body>', chart_html, re.DOTALL)
    chart = "".join(chart)
    chart_3 = chart

    df_site = df_site.rename({'5_인원(누계)': '누계인원'}, axis=1)
    df_site = df_site.reset_index()
    fig = px.pie(df_site, values='누계인원', names='현장명', color_discrete_sequence=px.colors.sequential.RdBu)
    fig.update_traces(textposition='inside', textinfo='percent+label')
    fig.update_layout(showlegend=False)
    chart_html = fig.to_html(include_plotlyjs='cdn')
    chart = re.findall('<body>(.*?)</body>', chart_html, re.DOTALL)
    chart = "".join(chart)
    chart_2 = chart

    context = {
        "title": "■ 본부별 인원 현황",
        "data": data_all,
        "title3": "■ 국내해외별 인원 현황",
        "data3": data_all_3,
        "title2": "■ 현장별 인원 현황",
        "data2": data_all_2,

        "title_chart1": "■ 본부별 그래프",
        "chart1": chart_1,
        "title_chart3": "■ 국내해외 그래프",
        "chart3": chart_3,
        "title_chart2": "■ 현장별 그래프",
        "chart2": chart_2,

        "date": datetime.today(),
    }
    return render(request, 'JobReport/JobReportHome.html', context)

def view_JobReportMan(request):
    template = 'JobReport/JobReportMan.html'
    dataControl = DC_JobReportMan()

    sDate_F = ''
    sDate_T = ''

    if request.method == 'GET':
        # try:
        #     sSecureKey = request.GET['SecureKey']
        # except:
        #     sSecureKey = ''
        try:    sBonbu = request.GET['bonbu']
        except: sBonbu = '전체'
        try:    sSite = request.GET['site']
        except: sSite = '전체'
        sDate_F = datetime.now().strftime('%Y-%m-') + '01'
        sDate_T = datetime.now().strftime('%Y-%m-%d')
    elif request.method == 'POST':
        # sSecureKey = request.POST['SecureKey']
        sBonbu = request.POST['bonbu']
        sSite = request.POST['site']
        sDate_F = request.POST['date_F']
        sDate_T = request.POST['date_T']
        try:
            sDate_F = datetime.strptime(sDate_F,'%Y-%m-%d').strftime('%Y-%m-%d')
            sDate_T = datetime.strptime(sDate_T,'%Y-%m-%d').strftime('%Y-%m-%d')
        except:
            sDate_F = datetime.now().strftime('%Y-%m-') + '01'
            sDate_T = datetime.now().strftime('%Y-%m-%d')

    df_data = dataControl.getJobReportManStatus(sBonbu,sSite,sDate_F,sDate_T)
    df_site = dataControl.getBonbuSiteList(sBonbu)
    # df = df[:100]

    columns = list(df_data.columns.values)
    colcount = len(columns)

    full_df = df_data.reset_index()
    old_Site = ''; old_Subcon = ''; old_BigGong = ''; old_Gong = ''
    lstDuplicate = []
    for row in full_df.values:
        if row[0] == old_Site:  duple = '0'
        else: duple = '1'

        if row[0] == old_Site and row[1] == old_Subcon:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong and row[3] == old_Gong:  duple = duple + '0'
        else: duple = duple + '1'

        old_Site    = row[0]
        old_Subcon  = row[1]
        old_BigGong = row[2]
        old_Gong    = row[3]

        lstDuplicate = lstDuplicate + [[row[0],row[1],row[2],row[3],duple]]

    df_Duples = pd.DataFrame(lstDuplicate)
    if len(df_Duples.values) > 0:
        full_df['Repeated'] = df_Duples[4]
    else:
        full_df['Repeated'] = '1111'

    json_data = full_df.to_json(orient='records')
    data = json.loads(json_data)

    json_sites = df_site.to_json(orient='records')
    sites = json.loads(json_sites)

    context = {
        'data': data,
        'sites': sites,
        'columns': columns,
        'colcount': colcount,
        'Bonbu': sBonbu,
        'Site': sSite,
        'F_Date': sDate_F,
        'T_Date': sDate_T,
        # 'SecureKey': sSecureKey,
    }

    return render(request,template,context)

def view_JobReportManExcel(request):
    sBonbu = request.GET['bonbu']
    sSite = request.GET['site']
    sDate_F = request.GET['date_F']
    sDate_T = request.GET['date_T']

    dataControl = DC_JobReportMan()
    df = dataControl.getJobReportManStatus(sBonbu,sSite,sDate_F,sDate_T)
    # df = df[:100]
    df = df.reset_index()
    df = df.fillna('')

    columns = list(df.columns.values.tolist())

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("sheet1")

    row_num = 0

    # 헤더 Write
    cell_format_head = workbook.add_format()
    cell_format_head.set_bold()
    cell_format_head.set_align('center')
    cell_format_head.set_bg_color('black')
    cell_format_head.set_font_color('white')
    for col_num in range(len(columns)):
        worksheet.write(row_num, col_num, columns[col_num], cell_format_head)

    # 데이터 Write
    cell_format_data = workbook.add_format()
    cell_format_data.set_font_color('black')
    for my_row in df.values:
        row_num = row_num + 1
        for col_num in range(len(columns)):
            sText = my_row[col_num]
            worksheet.write(row_num, col_num, sText,cell_format_data)

    workbook.close()
    output.seek(0)

    filename = 'JobReport[Man]_' + datetime.now().strftime('%Y%m%d_%H%M%S') +'.xlsx'
    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=%s' % filename

    return response

def view_JobReportEqu(request):
    template = 'JobReport/JobReportEqu.html'
    dataControl = DC_JobReportEqu()

    sDate_F = ''
    sDate_T = ''

    if request.method == 'GET':
        # try:
        #     sSecureKey = request.GET['SecureKey']
        # except:
        #     sSecureKey = ''
        try:    sBonbu = request.GET['bonbu']
        except: sBonbu = '전체'
        try:    sSite = request.GET['site']
        except: sSite = '전체'
        sDate_F = datetime.now().strftime('%Y-%m-') + '01'
        sDate_T = datetime.now().strftime('%Y-%m-%d')
    elif request.method == 'POST':
        # sSecureKey = request.POST['SecureKey']
        sBonbu = request.POST['bonbu']
        sSite = request.POST['site']
        sDate_F = request.POST['date_F']
        sDate_T = request.POST['date_T']
        try:
            sDate_F = datetime.strptime(sDate_F,'%Y-%m-%d').strftime('%Y-%m-%d')
            sDate_T = datetime.strptime(sDate_T,'%Y-%m-%d').strftime('%Y-%m-%d')
        except:
            sDate_F = datetime.now().strftime('%Y-%m-') + '01'
            sDate_T = datetime.now().strftime('%Y-%m-%d')

    df_data = dataControl.getJobReportEquStatus(sBonbu,sSite,sDate_F,sDate_T)
    df_site = dataControl.getBonbuSiteList(sBonbu)
    # df = df[:100]

    columns = list(df_data.columns.values)
    colcount = len(columns)

    full_df = df_data.reset_index()
    old_Site = ''; old_Subcon = ''; old_BigGong = ''; old_Gong = ''
    lstDuplicate = []
    for row in full_df.values:
        if row[0] == old_Site:  duple = '0'
        else: duple = '1'

        if row[0] == old_Site and row[1] == old_Subcon:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong and row[3] == old_Gong:  duple = duple + '0'
        else: duple = duple + '1'

        old_Site    = row[0]
        old_Subcon  = row[1]
        old_BigGong = row[2]
        old_Gong    = row[3]

        lstDuplicate = lstDuplicate + [[row[0],row[1],row[2],row[3],duple]]

    df_Duples = pd.DataFrame(lstDuplicate)
    if len(df_Duples.values) > 0:
        full_df['Repeated'] = df_Duples[4]
    else:
        full_df['Repeated'] = '1111'

    json_data = full_df.to_json(orient='records')
    data = json.loads(json_data)

    json_sites = df_site.to_json(orient='records')
    sites = json.loads(json_sites)


    context = {
        'data': data,
        'sites' : sites,
        'columns': columns,
        'colcount': colcount,
        'Bonbu': sBonbu,
        'Site': sSite,
        'F_Date': sDate_F,
        'T_Date': sDate_T,
        # 'SecureKey': sSecureKey,
    }

    return render(request,template,context)

def view_JobReportEquExcel(request):
    sBonbu = request.GET['bonbu']
    sSite = request.GET['site']
    sDate_F = request.GET['date_F']
    sDate_T = request.GET['date_T']

    dataControl = DC_JobReportEqu()
    df = dataControl.getJobReportEquStatus(sBonbu,sSite,sDate_F,sDate_T)
    # df = df[:100]
    df = df.reset_index()
    df = df.fillna('')

    columns = list(df.columns.values.tolist())

    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("sheet1")

    row_num = 0

    # 헤더 Write
    cell_format_head = workbook.add_format()
    cell_format_head.set_bold()
    cell_format_head.set_align('center')
    cell_format_head.set_bg_color('black')
    cell_format_head.set_font_color('white')
    for col_num in range(len(columns)):
        worksheet.write(row_num, col_num, columns[col_num], cell_format_head)

    # 데이터 Write
    cell_format_data = workbook.add_format()
    cell_format_data.set_font_color('black')
    for my_row in df.values:
        row_num = row_num + 1
        for col_num in range(len(columns)):
            sText = my_row[col_num]
            worksheet.write(row_num, col_num, sText,cell_format_data)

    workbook.close()
    output.seek(0)

    filename = 'JobReport[Equ]_' + datetime.now().strftime('%Y%m%d_%H%M%S') +'.xlsx'
    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=%s' % filename

    return response

def view_JobReportQnt(request):
    template = 'JobReport/JobReportQnt.html'
    dataControl = DC_JobReportQnt()

    sDate_F = ''
    sDate_T = ''

    if request.method == 'GET':
        # try:
        #     sSecureKey = request.GET['SecureKey']
        # except:
        #     sSecureKey = ''
        try:    sBonbu = request.GET['bonbu']
        except: sBonbu = '전체'
        try:    sSite = request.GET['site']
        except: sSite = '전체'
        sDate_F = datetime.now().strftime('%Y-%m-') + '01'
        sDate_T = datetime.now().strftime('%Y-%m-%d')
    elif request.method == 'POST':
        # sSecureKey = request.POST['SecureKey']
        sBonbu = request.POST['bonbu']
        sSite = request.POST['site']
        sDate_F = request.POST['date_F']
        sDate_T = request.POST['date_T']
        try:
            sDate_F = datetime.strptime(sDate_F,'%Y-%m-%d').strftime('%Y-%m-%d')
            sDate_T = datetime.strptime(sDate_T,'%Y-%m-%d').strftime('%Y-%m-%d')
        except:
            sDate_F = datetime.now().strftime('%Y-%m-') + '01'
            sDate_T = datetime.now().strftime('%Y-%m-%d')

    df_data = dataControl.getJobReportQntStatus(sBonbu,sSite,sDate_F,sDate_T)
    df_site = dataControl.getBonbuSiteList(sBonbu)
    # df = df[:100]

    columns = list(df_data.columns.values)
    colcount = len(columns)

    full_df = df_data.reset_index()
    old_Site = ''; old_Subcon = ''; old_BigGong = ''; old_Gong = ''; old_Building = ''; old_Location = ''
    lstDuplicate = []
    for row in full_df.values:
        if row[0] == old_Site:  duple = '0'
        else: duple = '1'

        if row[0] == old_Site and row[1] == old_Subcon:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong and row[3] == old_Gong:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong and row[3] == old_Gong and row[4] == old_Building:  duple = duple + '0'
        else: duple = duple + '1'

        if row[0] == old_Site and row[1] == old_Subcon and row[2] == old_BigGong and row[3] == old_Gong and row[4] == old_Building and row[5] == old_Location:  duple = duple + '0'
        else: duple = duple + '1'

        old_Site    = row[0]
        old_Subcon  = row[1]
        old_BigGong = row[2]
        old_Gong    = row[3]
        old_Building = row[4]
        old_Location = row[5]

        lstDuplicate = lstDuplicate + [[row[0],row[1],row[2],row[3],row[4],row[5],duple]]

    df_Duples = pd.DataFrame(lstDuplicate)
    if len(df_Duples.values) > 0:
        full_df['Repeated'] = df_Duples[6]
    else:
        full_df['Repeated'] = '111111'

    json_data = full_df.to_json(orient='records')
    data = json.loads(json_data)

    json_sites = df_site.to_json(orient='records')
    sites = json.loads(json_sites)


    context = {
        'data': data,
        'sites' : sites,
        'columns': columns,
        'colcount': colcount,
        'Bonbu': sBonbu,
        'Site': sSite,
        'F_Date': sDate_F,
        'T_Date': sDate_T,
        # 'SecureKey': sSecureKey,
    }

    return render(request,template,context)

def view_JobReportQntExcel(request):
    sBonbu = request.GET['bonbu']
    sSite = request.GET['site']
    sDate_F = request.GET['date_F']
    sDate_T = request.GET['date_T']

    dataControl = DC_JobReportQnt()
    df = dataControl.getJobReportQntStatus(sBonbu,sSite,sDate_F,sDate_T)
    # df = df[:100]
    df = df.reset_index()
    df = df.fillna('')

    columns = list(df.columns.values.tolist())
    output = io.BytesIO()
    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("sheet1")

    row_num = 0

    # 헤더 Write
    cell_format_head = workbook.add_format()
    cell_format_head.set_bold()
    cell_format_head.set_align('center')
    cell_format_head.set_bg_color('black')
    cell_format_head.set_font_color('white')
    for col_num in range(len(columns)):
        worksheet.write(row_num, col_num, columns[col_num], cell_format_head)

    # 데이터 Write
    cell_format_data = workbook.add_format()
    cell_format_data.set_font_color('black')
    for my_row in df.values:
        row_num = row_num + 1
        for col_num in range(len(columns)):
            sText = my_row[col_num]
            worksheet.write(row_num, col_num, sText,cell_format_data)


    workbook.close()
    output.seek(0)

    filename = 'JobReport[Qnt]_' + datetime.now().strftime('%Y%m%d_%H%M%S') +'.xlsx'
    response = HttpResponse(
        output,
        content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
    response['Content-Disposition'] = 'attachment; filename=%s' % filename

    return response

def view_BonbuSiteAjax(request):
    sBonbu = request.GET['bonbu']

    dataControl = DC_JobReportMan()
    df = dataControl.getBonbuSiteList(sBonbu)

    json_records = df.to_json(orient='records')
    data = json.loads(json_records)

    return JsonResponse(data,safe=False)

def view_ajax_test(request):
    if request.method == 'POST':
        data = json.loads(request.body)
        print(data)
        bonbu_code = data['bonbu_code']
        dom_code = data['dom_code']
        site_code = data['site_code']
        site_code = site_code.upper()
    else:
        print("error")

    domestic = ModelSubmissionStatus("DOMESTIC")
    df1 = domestic.resultBonsaExcel_Create()
    oversea = ModelSubmissionStatus("OVERSEA")
    df2 = oversea.resultBonsaExcel_Create()
    df = pd.concat([df1, df2])
    df = df.reset_index()
    df_all = df.copy()

    # 해당 현장이 있으면, 해당 현장의 데이터

    if bonbu_code == "토목사업본부":
        df3 = df[df['본부'].str.contains("토목사업본부", regex=False)]
        df4 = df[df['본부'].str.contains("Infrastructure Division", regex=False)]
        df = pd.concat([df3, df4])
    elif bonbu_code != "전체":
        df = df[df['본부'].str.contains(bonbu_code, regex=False)]
    if dom_code != "전체":
        df = df[df['국내해외'] == dom_code]
    if site_code != "전체":
        df = df[df['현장명'].str.contains(site_code, regex=False)]

    # if dataframe is empty, show all data
    if len(df.index) > 0:
        is_exist = True
    else:
        is_exist = False
        df = df_all
        site_code = ""
        dom_code = "전체"
        bonbu_code = "전체"


    df = df[['현장명','업체','1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                  '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']]
    df = df.set_index(['현장명', '업체'])

    df = pd.concat(
        [d.append(d.sum().rename(('', 'Total'))) for k, d in df.groupby(level=0)]).append(
        df.sum().rename(('Grand', 'Total')))

    # format numeric data
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    for col_name in numeric_columns:
        df[col_name] = df[col_name].apply(lambda x: "{:,.0f}".format(x))
    df = df.replace("0", "-")
    json_records = df.reset_index().to_json(orient='records')
    df_data = json.loads(json_records)
    context = {
        'requested': data,
        'df': df_data,
        'is_exist': is_exist,
    }

    return JsonResponse(context)


def view_daily_report_summary2(request):
    domestic = ModelSubmissionStatus("DOMESTIC")
    df1 = domestic.resultBonsaExcel_Create()
    oversea = ModelSubmissionStatus("OVERSEA")
    df2 = oversea.resultBonsaExcel_Create()
    df = pd.concat([df1, df2])
    df = df.reset_index()
    df_all = df

    if request.method == 'GET':

        df = df[['현장명','업체','1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']]
        df = df.set_index(['현장명', '업체'])

        df = pd.concat(
            [d.append(d.sum().rename(('', 'Total'))) for k, d in df.groupby(level=0)]).append(
            df.sum().rename(('Grand', 'Total')))

        is_exist = True
        site_code = "전체"
        dom_code = "전체"
        bonbu_code = "전체"

    elif request.method == 'POST':

        site_code = request.POST.get('site_code', "전체")
        site_code = site_code.upper()
        print("현장코드 : ", site_code)
        dom_code = request.POST.get('dom_code', "전체")
        print("국내/해외 : ", dom_code)
        bonbu_code = request.POST.get('bonbu_code', "전체")
        print("본부 : ", bonbu_code)

        # 해당 현장이 있으면, 해당 현장의 데이터

        if bonbu_code == "토목사업본부":
            df3 = df[df['본부'].str.contains("토목사업본부", regex=False)]
            df4 = df[df['본부'].str.contains("Infrastructure Division", regex=False)]
            df = pd.concat([df3, df4])
        elif bonbu_code != "전체":
            df = df[df['본부'].str.contains(bonbu_code, regex=False)]
        if dom_code != "전체":
            df = df[df['국내해외'] == dom_code]
        if site_code != "전체":
            df = df[df['현장명'].str.contains(site_code, regex=False)]

        # if dataframe is empty, show all data
        if len(df.index) > 0:
            is_exist = True
        else:
            is_exist = False
            df = df_all
            site_code = "전체"
            dom_code = "전체"
            bonbu_code = "전체"


        # 본부, 국내해외 제외하고 데이터를 보여줌
        df = df[['현장명','업체','1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']]
        df = df.set_index(['현장명', '업체'])

        df = pd.concat(
            [d.append(d.sum().rename(('', 'Total'))) for k, d in df.groupby(level=0)]).append(
            df.sum().rename(('Grand', 'Total')))

    # format numeric data
    numeric_columns = ['1_인원(D3)', '2_인원(D2)', '3_인원(D1)', '4_인원(금일)', '5_인원(누계)',
                      '6_장비(D3)', '7_장비(D2)', '8_장비(D1)', '9_장비(금일)', 'A_장비(누계)']
    for col_name in numeric_columns:
        df[col_name] = df[col_name].apply(lambda x: "{:,.0f}".format(x))
    df = df.replace("0", "-")

    json_records = df.reset_index().to_json(orient='records')
    data_all = json.loads(json_records)

    context = {
        "title": "■ 일일 제출 현황",
        "date": datetime.today(),
        "data": data_all,
        "site_code": site_code,
        "bonbu_code": bonbu_code,
        "dom_code": dom_code,
        "is_exist": is_exist,
    }
    return render(request, 'JobReport/JobReportSummaryAjax.html', context)


def view_predict_period(request):
    dataControl = DataControl()
    sSql = f"select max(년월) from predict_period where 년월 ~ '^[0-9\.]+$'"
    column_names = ["년월"]
    df_key_month = dataControl.DF_from_DB_with_sSql_colName(sSql, column_names)
    key_month = df_key_month["년월"][0]

    if request.method == 'GET':
        bonbu_code = "전체"
        siteName_code = "전체"
        gongjong_code = "레미콘타설"

    elif request.method == 'POST':
        # find out all requested data
        bonbu_code = request.POST.get('bonbu_code', "전체")
        print("본부 : ", bonbu_code)
        siteName_code = request.POST.get('siteName_code', "전체")
        print("현장명 : ", siteName_code)
        gongjong_code = request.POST.get('gongjong_code', "전체")
        print("공종코드 : ", gongjong_code)
        # key_month = request.POST.get('key_month')
        # print("해당월 : ", key_month)

    modelPredictPeriod = ModelPredictPeriod(key_month)
    df, bonbu_code, siteName_code, gongjong_code, bonbu_list, siteName_list, gongjong_list, M5, M4, M3, M2, M1, M0\
        = modelPredictPeriod.model_predict_period(bonbu_code, siteName_code, gongjong_code)

    # numeric data format
    df['계약물량'] = df['계약물량'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['전월누계물량'] = df['전월누계물량'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['금월물량'] = df['금월물량'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['금월누계'] = df['금월누계'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['잔여량'] = df['잔여량'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['진행율'] = df['진행율'].apply(lambda x: '{:.0%}'.format(x/100) if x != 0 else "-")
    df['월평균소화물량'] = df['월평균소화물량'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df[f'{M5}'] = df[f'{M5}'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df[f'{M4}'] = df[f'{M4}'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df[f'{M3}'] = df[f'{M3}'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df[f'{M2}'] = df[f'{M2}'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df[f'{M1}'] = df[f'{M1}'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df[f'{M0}'] = df[f'{M0}'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['잔여개월수'] = df['잔여개월수'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['잔여공기예측_누적평균기준'] = df['잔여공기예측_누적평균기준'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['잔여공기예측_최근월기준'] = df['잔여공기예측_최근월기준'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")
    df['잔여공기예측_3월평균기준'] = df['잔여공기예측_3월평균기준'].apply(lambda x: '{:,.0f}'.format(x) if x != 0 else "-")

    # make response
    json_records = df.to_json(orient='records')
    df_data = json.loads(json_records)
    context = {
        'data': df_data,
        'bonbu_code': bonbu_code,
        'siteName_code': siteName_code,
        'gongjong_code': gongjong_code,
        'bonbu_list': bonbu_list,
        'siteName_list': siteName_list,
        'gongjong_list': gongjong_list,
        'key_month': key_month,
        'M5': M5, 'M4': M4, 'M3': M3, 'M2': M2, 'M1': M1, 'M0': M0,
    }

    return render(request, 'JobReport/JobReportPredict.html', context)


def view_predict_period_excel_download(request):
    # find out all requested data
    bonbu_code = request.POST.get('bonbu_code', "전체")
    siteName_code = request.POST.get('siteName_code', "전체")
    gongjong_code = request.POST.get('gongjong_code', "전체")
    key_month = request.POST.get('key_month')
    # operation data
    modelPredictPeriod = ModelPredictPeriod(key_month)
    df, bonbu_code, siteName_code, gongjong_code, bonbu_list, site_list, gongjong_list, M5, M4, M3, M2, M1, M0 \
        = modelPredictPeriod.model_predict_period(bonbu_code, siteName_code, gongjong_code)


    # numeric data format
    df['계약물량'] = df['계약물량'].apply(lambda x: int(round(x, 2)))
    df['전월누계물량'] = df['전월누계물량'].apply(lambda x: int(round(x, 2)))
    df['금월물량'] = df['금월물량'].apply(lambda x: int(round(x, 2)))
    df['금월누계'] = df['금월누계'].apply(lambda x: int(round(x, 2)))
    df['잔여량'] = df['잔여량'].apply(lambda x: int(round(x, 2)))
    df['진행율'] = df['진행율'].apply(lambda x: int(round(x, 2)))
    df['월평균소화물량'] = df['월평균소화물량'].apply(lambda x: int(round(x, 2)))
    df[f'{M5}'] = df[f'{M5}'].apply(lambda x: int(round(x, 2)))
    df[f'{M4}'] = df[f'{M4}'].apply(lambda x: int(round(x, 2)))
    df[f'{M3}'] = df[f'{M3}'].apply(lambda x: int(round(x, 2)))
    df[f'{M2}'] = df[f'{M2}'].apply(lambda x: int(round(x, 2)))
    df[f'{M1}'] = df[f'{M1}'].apply(lambda x: int(round(x, 2)))
    df[f'{M0}'] = df[f'{M0}'].apply(lambda x: int(round(x, 2)))
    df['잔여개월수'] = df['잔여개월수'].apply(lambda x: int(round(x, 2)))
    df['잔여공기예측_누적평균기준'] = df['잔여공기예측_누적평균기준'].apply(lambda x: int(round(x, 2)))
    df['잔여공기예측_최근월기준'] = df['잔여공기예측_최근월기준'].apply(lambda x: int(round(x, 2)))
    df['잔여공기예측_3월평균기준'] = df['잔여공기예측_3월평균기준'].apply(lambda x: int(round(x, 2)))

    # 특수문자 제거
    import re
    gongjong_code_re = re.sub('\W+', '', gongjong_code)
    # \W  : Matches any character which is not a word character.
    # This is the opposite of \w. If the ASCII flag is used this becomes the equivalent of [^a-zA-Z0-9_].
    # If the LOCALE flag is used, matches characters which are neither alphanumeric
    # in the current locale nor the underscore.
    # +   : 1번이상 반복

    # save in memory by pandas / io
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer) as writer:
        df.to_excel(writer, sheet_name=gongjong_code_re)
    buffer.seek(0)

    # response
    now = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"predictPeriod_{key_month}_{now}.xlsx"
    response = HttpResponse(buffer.read(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = f"attachment; filename={filename}"

    # memory close
    buffer.close()

    return response

def view_predict_period_ajax(request):
    if request.method == 'POST':
        requested_data = json.loads(request.body)
        print("requested_data : ", requested_data)
        if 'bonbu_code' in requested_data:
            bonbu_code = requested_data['bonbu_code']
        else:
            bonbu_code = "전체"
        if 'siteName_code' in requested_data:
            siteName_code = requested_data['siteName_code']
        else:
            siteName_code = "전체"
        if 'gongjong_code' in requested_data:
            gongjong_code = requested_data['gongjong_code']
        else:
            gongjong_code = "전체"
        key_month = requested_data['key_month']
    else:
        print("요청 방식이 POST가 아닙니다.")

    # operation data
    modelPredictPeriod = ModelPredictPeriod(key_month)
    siteName_list, gongjong_list \
        = modelPredictPeriod.model_predict_period_ajax(bonbu_code, siteName_code)

    # if dataframe is empty, show all data
    if len(siteName_list) > 1 or len(gongjong_list) > 1:
        is_exist = True
    else:
        is_exist = False

    # response
    context = {
        'requested': requested_data,
        'siteName_list': siteName_list,
        'gongjong_list': gongjong_list,
        'is_exist': is_exist,
    }

    return JsonResponse(context)

def view_JobReportPlnt(request):
    if request.method == 'GET':
        projectcode = None
        discipline = None
        subcontractor = None
        category = None
        cwa = None
        cwp = None
        iwp = None
        # f_date = "2022-03-01"
        f_date = f"{datetime.now().strftime('%Y-%m')}-01"
        t_date = datetime.now().strftime('%Y-%m-%d')

    if request.method == 'POST':
        print(request.POST)
        projectcode = request.POST.get('projectcode', None)
        discipline = request.POST.get('discipline', None)
        subcontractor = request.POST.get('subcontractor', None)
        category = request.POST.get('category', None)
        cwa = request.POST.get('cwa', None)
        cwp = request.POST.get('cwp', None)
        iwp = request.POST.get('iwp', None)
        f_date = request.POST.get('f_date', "2022-03-01")
        t_date = request.POST.get('t_date', datetime.now().strftime('%Y-%m-%d'))

        # 전체 -> None
        projectcode = projectcode if projectcode != '전체' else None
        discipline = discipline if discipline != '전체' else None
        subcontractor = subcontractor if subcontractor != '전체' else None
        category = category if category != '전체' else None
        cwa = cwa if cwa != '전체' else None
        cwp = cwp if cwp != '전체' else None
        iwp = iwp if iwp != '전체' else None


    dc_plnt = DC_JobReportPlnt()
    df = dc_plnt.export_plnt_data(f_date, t_date, projectcode, discipline, subcontractor, category, cwa, cwp, iwp)
    if df.empty:
        return render(request, 'JobReport/JobReportPlnt.html', context={'colcount': 0})

    columns = list(df.columns.values)
    colcount = len(columns)

    # df data
    df = df.reset_index()
    json_records = df.to_json(orient='records')
    df_data = json.loads(json_records)

    # df index
    df_index = dc_plnt.df_index_arranger(df)
    json_records = df_index.to_json(orient='records')
    df_index = json.loads(json_records)

    # search parameter
    projectcodes = pd.unique(df.projectcode)
    disciplines = pd.unique(df.discipline)
    subcontractors = pd.unique(df.subcontractor)
    categories = pd.unique(df.category)
    cwas = pd.unique(df.cwa)
    cwps = pd.unique(df.cwp)
    iwps = pd.unique(df.iwp)

    projectcodes = [i for i in projectcodes if i not in ['', ' ', '-'] and 'Total' not in i]
    disciplines = [i for i in disciplines if i not in ['', ' ', '-'] and 'Total' not in i]
    subcontractors = [i for i in subcontractors if i not in ['', ' ', '-'] and 'Total' not in i]
    categories = [i for i in categories if i not in ['', ' ', '-'] and 'Total' not in i]
    cwas = [i for i in cwas if i not in ['', ' ', '-'] and 'Total' not in i]
    cwps = [i for i in cwps if i not in ['', ' ', '-'] and 'Total' not in i]
    iwps = [i for i in iwps if i not in ['', ' ', '-'] and 'Total' not in i]

    # None -> 전체
    projectcode = '전체' if projectcode is None else projectcode
    discipline = '전체' if discipline is None else discipline
    subcontractor = '전체' if subcontractor is None else subcontractor
    category = '전체' if category is None else category
    cwa = '전체' if cwa is None else cwa
    cwp = '전체' if cwp is None else cwp
    iwp = '전체' if iwp is None else iwp

    # make response
    context = {
        'data': df_data,
        'index': df_index,
        'columns': columns,
        'colcount': colcount,

        'projectcode': projectcode,
        'discipline': discipline,
        'subcontractor': subcontractor,
        'category': category,
        'cwa': cwa,
        'cwp': cwp,
        'iwp': iwp,

        'projectcodes': projectcodes,
        'disciplines': disciplines,
        'subcontractors': subcontractors,
        'categories': categories,
        'cwas': cwas,
        'cwps': cwps,
        'iwps': iwps,
        'f_date': f_date,
        't_date': t_date,
    }
    return render(request, 'JobReport/JobReportPlnt.html', context)

def view_JobReportPlntExcel(request):
    print(request.POST)
    projectcode = request.POST.get('projectcode', None)
    discipline = request.POST.get('discipline', None)
    subcontractor = request.POST.get('subcontractor', None)
    category = request.POST.get('category', None)
    cwa = request.POST.get('cwa', None)
    cwp = request.POST.get('cwp', None)
    iwp = request.POST.get('iwp', None)
    f_date = request.POST.get('f_date', "2022-03-01")
    t_date = request.POST.get('t_date', datetime.now().strftime('%Y-%m-%d'))

    # 전체 -> None
    projectcode = projectcode if projectcode != '전체' else None
    discipline = discipline if discipline != '전체' else None
    subcontractor = subcontractor if subcontractor != '전체' else None
    category = category if category != '전체' else None
    cwa = cwa if cwa != '전체' else None
    cwp = cwp if cwp != '전체' else None
    iwp = iwp if iwp != '전체' else None

    # call data
    dc_plnt = DC_JobReportPlnt()
    df = dc_plnt.export_plnt_data(f_date, t_date, projectcode, discipline, subcontractor, category, cwa, cwp, iwp)

    # # make directory
    # import os
    # savePath = f"D:\Project_Data\JobReport\Files_PLNT\{datetime.now().strftime('%Y-%m-%d')}"
    # os.makedirs(savePath, exist_ok=True)
    # now = datetime.now().strftime("%Y%m%d_%H%M_%S")
    # filename = f"jobReportPlnt_{now}.xlsx"
    # outputFileName = savePath + "\\" + filename
    # if os.path.exists(outputFileName):
    #     os.remove(outputFileName)
    #
    # # save file
    # with pd.ExcelWriter(outputFileName) as writer:
    #     df.to_excel(writer, sheet_name="data")
    #
    # # style excel format
    # if not df.empty:
    #     dc_plnt.excel_styler_plnt_prd_cum(outputFileName)
    #
    # # file response
    # response = FileResponse(open(outputFileName, 'rb'))
    # return response


# -----# -----
    # save in memory by pandas / io
    import io
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer) as writer:
        df.to_excel(writer, sheet_name='data')
    buffer.seek(0)

    # response
    now = datetime.now().strftime("%Y%m%d%H%M%S")
    filename = f"jobReportPlnt_{now}.xlsx"
    response = HttpResponse(buffer.read(),
                            content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response['Content-Disposition'] = f"attachment; filename={filename}"

    # memory close
    buffer.close()

    return response

# # -----# -----
#     columns = list(df.columns.values.tolist())
#     output = io.BytesIO()
#
#     workbook = xlsxwriter.Workbook(buffer)
#     worksheet = workbook.add_worksheet("sheet1")
#
#     row_num = 0
#
#     # 헤더 Write
#     cell_format_head = workbook.add_format()
#     cell_format_head.set_bold()
#     cell_format_head.set_align('center')
#     cell_format_head.set_bg_color('black')
#     cell_format_head.set_font_color('white')
#     for col_num in range(len(columns)):
#         worksheet.write(row_num, col_num, columns[col_num], cell_format_head)
#
#     # 데이터 Write
#     cell_format_data = workbook.add_format()
#     cell_format_data.set_font_color('black')
#     for my_row in df.values:
#         row_num = row_num + 1
#         for col_num in range(len(columns)):
#             sText = my_row[col_num]
#             worksheet.write(row_num, col_num, sText, cell_format_data)
#
#     workbook.close()
#     output.seek(0)
#
#     filename = 'JobReport[Qnt]_' + datetime.now().strftime('%Y%m%d_%H%M%S') +'.xlsx'
#     response = HttpResponse(
#         output,
#         content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
#     )
#     response['Content-Disposition'] = 'attachment; filename=%s' % filename
#
#     return response
#
# # -----



def view_JobReportPlntAjax(request):
    requested_data = json.loads(request.body)
    print("requested_data : ", requested_data)

    projectcode = requested_data['projectcode']
    discipline = requested_data['discipline']
    subcontractor = requested_data['subcontractor']
    category = requested_data['category']
    cwa = requested_data['cwa']
    cwp = requested_data['cwp']
    iwp = requested_data['iwp']
    f_date = requested_data['f_date']
    t_date = requested_data['t_date']

    projectcode = None if projectcode == '전체' else projectcode
    discipline = None if discipline == '전체' else discipline
    subcontractor = None if subcontractor == '전체' else subcontractor
    category = None if category == '전체' else category
    cwa = None if cwa == '전체' else cwa
    cwp = None if cwp == '전체' else cwp
    iwp = None if iwp == '전체' else iwp

    # operation data
    # call data
    dc_plnt = DC_JobReportPlnt()
    df = dc_plnt.export_plnt_data(f_date, t_date, projectcode, discipline, subcontractor, category, cwa, cwp, iwp)
    df = df.reset_index()
    df = df[["projectcode", "discipline", "subcontractor", "category", "cwa", "cwp", "iwp"]]

    # 데이터가 없는 경우

    # 전체데이터의 경우
    if projectcode == discipline == subcontractor == category == cwa == cwp == iwp == None:
        projectcodes, disciplines, subcontractors, categories, cwas, cwps, iwps = dc_plnt.select_parameter(df)
    else:
        for key, value in requested_data.items():
            if key not in ['f_date', 't_date']:
                if value != '전체':
                    df = df[df[key] == value]
        projectcodes, disciplines, subcontractors, categories, cwas, cwps, iwps = dc_plnt.select_parameter(df)

    projectcode = projectcodes
    discipline = disciplines
    subcontractor = subcontractors
    category = categories
    cwa = cwas
    cwp = cwps
    iwp = iwps

    # response
    context = {
        'projectcode': projectcode,
        'discipline': discipline,
        'subcontractor': subcontractor,
        'category': category,
        'cwa': cwa,
        'cwp': cwp,
        'iwp': iwp,
    }
    print("context : ", context)
    return JsonResponse(context)
