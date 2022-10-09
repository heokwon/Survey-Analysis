import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

## bar_plot
### excel_path : 엑셀파일 경로
### column_name : 원하는 컬럼 이름
### figsize : plot size, default = (10.8), 크기 조정 가능
### bar_plot(엑셀경로, 분석하고자 하는 컬럼 이름, plt사이즈)
def bar_plot(excel_path, column_name, figsize=(10,8)):
    df = pd.read_excel(excel_path)
    df_column = pd.DataFrame(df[column_name].value_counts())
    df_column = df_column.rename(columns={column_name:'Count'})

    plt.figure(figsize=figsize)
    plt.xlabel('Count')
    plt.title(column_name)

    sns.barplot(x=df_column['Count'],y=df_column.index)

## 성별 & 연령대별 구내식당 이용빈도or구내식당 선호 메뉴 시각화
## gender_histplot(엑셀경로, 남or여, 연령대, 구내식당 이용빈도or구내식당 선호 메뉴)
def gender_histplot(excel_path, gender,MorF, age_group, column_name, figsize=(10,8)):
    df = pd.read_excel(excel_path)
    df = df[df[gender]==MorF]
    plt.figure(figsize=figsize)
    plt.title(gender+' '+age_group+' '+column_name)

    sns.histplot(data=df, y=age_group, hue=column_name, multiple="stack")

## strip_plot(엑셀 경로, 성별or연령대, 구내식당 이용빈도or구내식당 선호 메뉴)
def strip_plot(excel,column,sec_column,figsize=(10,10)):
    df=pd.read_excel(excel)
    plt.figure(figsize=(10,10))
    plt.title(column+' '+sec_column)
    sns.stripplot(x=column, y=sec_column, data=df)

## pie_plot
### excel_path : 엑셀파일 경로
### column_name : 원하는 컬럼 이름
### figsize : plot size, default = (10.10), 크기 조정 가능
### pie_plot(엑셀경로, 분석하고자 하는 컬럼이름, plt사이즈)
def pie_plot(excel_path, column_name, figsize=(10,10)):

    df = pd.read_excel(excel_path)
    df_column = pd.DataFrame(df[column_name].value_counts())
    df_column = df_column.rename(columns={column_name:'Count'})

    ratio = [i for i in df_column["Count"]]
    ratio = [i/sum(ratio) for i in ratio]
    labels = [i for i in df_column.index]
    wedgeprops={'width': 0.6, 'edgecolor': 'w', 'linewidth': 5} # 부채꼴 스타일 지정하기

    plt.figure(figsize=figsize)
    plt.title(column_name)
    plt.pie(ratio, labels=labels,autopct='%.1f%%', wedgeprops=wedgeprops)
    plt.show

## 두개 컬럼을 사용하여 분석
## ex) first_column = "연령대",contents = "40대" ,sec_column = "구내식당 선호메뉴"
##     >> 연령대가 40대인 사원들의 구내식당 선호메뉴 비중
def columns_analysis(excel_path, first_column, contents, sec_column, figsize=(10,10)):
    df = pd.read_excel(excel_path)
    df = df[df[first_column]== contents]
    df = pd.DataFrame(df[sec_column].value_counts())
    df = df.rename(columns={sec_column:'Count'})

    ratio = [i for i in df["Count"]]
    ratio = [i/sum(ratio) for i in ratio]
    labels = [i for i in df.index]
    wedgeprops={'width': 0.6, 'edgecolor': 'w', 'linewidth': 5} # 부채꼴 스타일 지정하기

    plt.figure(figsize=figsize)
    plt.title(contents+' '+sec_column)
    plt.pie(ratio, labels=labels,autopct='%.1f%%', wedgeprops=wedgeprops)
    plt.show

## 성별 + 연령대 분석
## ex) 남 40대 구내식당 선호메뉴
## gender_analysis_pieplot(엑셀경로, 성별, 남or여, 20 30 40 50대 중 택1, 구내식당 이용빈도 or 구내식당 선호 메뉴)
def gender_analysis_pieplot(excel_path, gender_column, age_columns, Gender, age_group, sec_column,figsize=(10,10)):
    df = pd.read_excel(excel_path)
    gender = df[df[gender_column]==Gender]
    gender_AgeGroup = gender[gender[age_columns]==age_group]
    gender_AgeGroup = pd.DataFrame(gender_AgeGroup[sec_column].value_counts())

    ratio = [i for i in gender_AgeGroup[sec_column]]
    ratio = [i/sum(ratio) for i in ratio]
    labels = [i for i in gender_AgeGroup.index]
    wedgeprops={'width': 0.6, 'edgecolor': 'w', 'linewidth': 5} 

    plt.figure(figsize=figsize)
    plt.title(Gender+' '+age_group+' '+sec_column)
    plt.pie(ratio, labels=labels,autopct='%.1f%%', wedgeprops=wedgeprops)
    plt.show
    
## 성별&연령대별 구내식당 선호 메뉴 or 구내식당 이용빈도 카운트
## count_df(엑셀경로, 연령대. 구내식당 선호 메뉴or구내식당 이용빈도, 성별,남or여)
def count_df(excel_path, age_col, analysis_col, gen_col, gender):
    df = pd.read_excel(excel_path)
    col_list = [i for i in df[age_col].drop_duplicates()]
    index_list = [i for i in df[analysis_col].drop_duplicates()]
    
    new_df = {}
    for i in range(len(col_list)):
        new_df[col_list[i]] = []

    for col in col_list:
        for index in index_list:
            new = (df[gen_col]==gender) & (df[age_col]==col) & (df[analysis_col]==index)
            count = len(df[new])
            new_df[col].append(count)

    new_df = pd.DataFrame(new_df,index = index_list)
    new_df['합계'] = new_df.sum(axis=1)
    new_df.loc['합계'] = new_df.sum(axis=0)

    return new_df