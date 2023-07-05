import streamlit as st
import pandas as pd
from PIL import Image
from openpyxl.utils.dataframe import dataframe_to_rows



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
cross1 = ""
cross1 = st.sidebar.selectbox(
    "クロス軸",
    q_list
)
cross2 = st.sidebar.selectbox(
    "設問",
    q_list
)

def pivot_calc(cross1, cross2):
    cross1_key = str(cross1) + " "
    cross2_key = str(cross2) + " "

    df_mt = pd.DataFrame()
    #ピヴォットの計算
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
    
    return df_mt

df_mt = pivot_calc(cross1, cross2)

if len(df_mt) != 0:
    df_mt_p = df_mt.copy()
    df_mt_p.iloc[:, :] = df_mt_p.iloc[:, :].div(df_mt_p.iloc[:, -1], axis=0)
else:
    df_mt_p = pd.DataFrame()

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
text = ""
out_cross = []

expander2 = st.sidebar.expander('ダウロード')
if len(q_list) != 0:
    text = expander2.text_input("出力ファイル名")
    file_name = text + '.csv'
    #out_cross = expander2.multiselect(label="保存する設問  （※クロス軸：" + cross1 + "）", options=q_list)
    out_cross = expander2.multiselect(label="保存する設問(※クロス軸：" + cross1 +")", options=q_list)

    #保存用データ格納
    append_pivot =[]

    if expander2.button("データ作成"):
        if text != "":
            for oc in out_cross:
                out_df_mt = pivot_calc(cross1, oc)
                out_df_mt_p = out_df_mt.copy()
                out_df_mt_p.iloc[:, :] = out_df_mt_p.iloc[:, :].div(out_df_mt_p.iloc[:, -1], axis=0)

                # 空行を追加
                append_pivot.append([oc,'=========================================================================================='])

                # 空行を追加
                append_pivot.append(['▼人数▼',])
                # DataFrame1をExcelシートに書き込む
                for row in dataframe_to_rows(out_df_mt, index=True, header=True):
                    append_pivot.append(row)
                
                # 空行を追加
                append_pivot.append([])

                # 空行を追加
                append_pivot.append(['▼割合▼',])
                # DataFrame2をExcelシートに書き込む
                for row in dataframe_to_rows(out_df_mt_p, index=True, header=True):
                    append_pivot.append(row)

                # 空行を追加
                append_pivot.append([])

            
            csv = pd.DataFrame(append_pivot).to_csv(index=False, header=False).encode('ms932')
            expander2.download_button(  label='Data Download', 
                                        data=csv, 
                                        file_name=text + '.csv',
                                        mime='text/csv'
                                        )
        else:
            expander2.markdown('<span style="color:red">※ファイル名を入力してください</span>', unsafe_allow_html=True)
