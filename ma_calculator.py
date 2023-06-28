import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
#import plotly.graph_objects as go

#st.title('MA Calculator')

def ma():
    img = Image.open('header.png')
    st.image(img, use_column_width=True)

    uploaded_file = st.sidebar.file_uploader("ファイル選択", type="csv")

    df = pd.DataFrame()
    merged_table = pd.DataFrame()

    expander1 = st.expander('読込ファイル')
    if uploaded_file is not None:
        df = pd.DataFrame(pd.read_csv(uploaded_file, encoding='ms932'))
        expander1.write(df)
    else:
        expander1.write('CSVファイルをアップロードしてください')


    #クロス軸のプルダウン作成
    q_list=[]
    colmuns = list(df.columns)

    for column in colmuns:
        first_word = column.split()[0]
        if first_word[0] == "Q":
            if first_word not in q_list:
                q_list.append(first_word)
    cross1 = st.sidebar.selectbox(
        "クロス軸",
        q_list
    )
    cross2 = st.sidebar.selectbox(
        "設問",
        q_list
    )

    cross1_key = str(cross1) + " "
    cross2_key = str(cross2) + " "

    #ピヴォットの計算
    df_mt = pd.DataFrame()
    df_mt_p = pd.DataFrame()
    if cross1 != cross2:
        for c in colmuns:
            if c.startswith(cross1_key) == True:
                c1 = c

        pivot_table = []
        for cn in colmuns:
            if cn.startswith(cross2_key) == True and "その他テキスト" not in cn:
                pt = df.pivot_table(index=[c1], columns=[cn], values=colmuns[0], aggfunc='count')
                pivot_table.append(pt)

        if len(pivot_table) != 0:
            merged_table = pd.concat(pivot_table, axis=1)

        #母体の計算
        botai = df.pivot_table(index=[c1], values=colmuns[0], aggfunc='count')
        df_mt = pd.DataFrame(merged_table)
        df_mt['母体数'] = botai

        df_mt_p = df_mt.copy()
        df_mt_p.iloc[:, :] = df_mt_p.iloc[:, :].div(df_mt_p.iloc[:, -1], axis=0)

    #タブ作成
    tab1,tab2 = st.tabs(["ピボット","グラフ"])

    with tab1:
        #結果の表示
        st.info('人数')
        st.dataframe(df_mt)

        st.info('割合')
        st.dataframe(df_mt_p.style.format('{:.2%}'))

    with tab2:
        if len(df_mt_p) != 0:
            #グラフの種類選択
            stock = st.radio(label='グラフの種類選択',
                            options=('棒グラフ', '面グラフ', '折れ線グラフ'),
                            index=0,
                            horizontal=True)
            #グラフ用df用意
            df_mt_p_glaph = df_mt_p.drop('母体数', axis=1)
            if stock == "棒グラフ":
                st.bar_chart(df_mt_p_glaph.style.format('{:.2%}'))
            elif stock == "面グラフ":
                st.area_chart(df_mt_p_glaph.style.format('{:.2%}'))
            elif stock == "折れ線グラフ":
                st.line_chart(df_mt_p_glaph.style.format('{:.2%}'))

    #DL設定
    expander2 = st.sidebar.expander('ダウロード')
    text = expander2.text_input("出力ファイル名")
    button = expander2.button("保存")
    file_name = text + '.xlsx'

    if button:
        if text != "":
            try:
                with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    df_mt.to_excel(writer, sheet_name=cross2, index=True, header=True)
                with pd.ExcelWriter(file_name, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    df_mt_p.to_excel(writer, sheet_name=cross2, index=True, header=True, startrow=len(df_mt)+2)
                expander2.markdown('<span style="color:green">保存しました</span>', unsafe_allow_html=True)
            except FileNotFoundError:
                with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                    df_mt.to_excel(writer, sheet_name=cross2, index=True, header=True)
                    df_mt_p.to_excel(writer, sheet_name=cross2, index=True, header=True, startrow=len(df_mt)+2)
                expander2.markdown('<span style="color:green">保存しました</span>', unsafe_allow_html=True)
            except PermissionError:
                expander2.markdown('<span style="color:red">※指定ファイルを閉じてください</span>', unsafe_allow_html=True)
        else:
            expander2.markdown('<span style="color:red">※ファイル名を入力してください</span>', unsafe_allow_html=True)
if __name__ == "__main__": 
    ma()