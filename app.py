
import sqlite3
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side

# SQLite 데이터베이스와 연결하는 함수
def connect_db():
    return sqlite3.connect('restaurant_menu.db')

# SQLite 데이터베이스에 메뉴를 삽입하는 함수
def insert_menu(menu_name):
    conn = connect_db()
    c = conn.cursor()
    c.execute('INSERT INTO Menus (MenuName) VALUES (?)', (menu_name,))
    conn.commit()
    conn.close()

# SQLite 데이터베이스에 재료를 삽입하는 함수
def insert_ingredient(ingredient_name, price):
    conn = connect_db()
    c = conn.cursor()
    c.execute('INSERT INTO Ingredients (IngredientName, Price) VALUES (?, ?)', (ingredient_name, price))
    conn.commit()
    conn.close()

# SQLite 데이터베이스에서 메뉴와 재료를 조회하는 함수
def get_data():
    conn = connect_db()
    c = conn.cursor()
    c.execute('SELECT * FROM Menus')
    menus = c.fetchall()
    c.execute('SELECT * FROM Ingredients')
    ingredients = c.fetchall()
    conn.close()
    return menus, ingredients

def save_to_excel(materials_df, total_material_cost, total_labor_cost, total_cost):
    wb = Workbook()
    ws = wb.active
    
    # 헤더 작성
    headers = ['IngredientName', 'Price', 'Quantity', 'Total Cost (KRW)']
    ws.append(headers)
    
    # 데이터 작성
    for idx, row in materials_df.iterrows():
        ws.append([row['IngredientName'], row['Price'], row['Quantity'], row['TotalCost']])
    
    # 합계 작성
    ws.append([])
    ws.append(['Total Material Cost', total_material_cost])
    ws.append(['Total Labor Cost', total_labor_cost])
    ws.append(['Total Cost', total_cost])
    
    # 테두리 추가
    thin = Side(border_style="thin", color="000000")
    for row in ws.iter_rows():
        for cell in row:
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    
    # 엑셀 파일 저장
    excel_file_path = "견적서.xlsx"
    wb.save(excel_file_path)
    return excel_file_path

def main():
    st.set_page_config(page_title="레스토랑 메뉴 관리")
    st.title("레스토랑 메뉴 및 재료 관리 시스템")

    if "menu_df" not in st.session_state or "ingredient_df" not in st.session_state:
        menus, ingredients = get_data()
        st.session_state.menu_df = pd.DataFrame(menus, columns=["ID", "MenuName"])
        st.session_state.ingredient_df = pd.DataFrame(ingredients, columns=["ID", "IngredientName", "Price"])

    st.sidebar.header("메뉴 추가")
    with st.sidebar.form("메뉴 추가 양식"):
        menu_name = st.text_area("메뉴 이름")
        menu_submitted = st.form_submit_button("메뉴 추가")

    if menu_submitted:
        insert_menu(menu_name)
        st.sidebar.write("메뉴가 추가되었습니다!")
        menus, _ = get_data()
        st.session_state.menu_df = pd.DataFrame(menus, columns=["ID", "MenuName"])

    st.sidebar.header("재료 추가")
    with st.sidebar.form("재료 추가 양식"):
        ingredient_name = st.text_area("재료 이름")
        price = st.number_input("가격", min_value=0.0, step=0.1)
        ingredient_submitted = st.form_submit_button("재료 추가")

    if ingredient_submitted:
        insert_ingredient(ingredient_name, price)
        st.sidebar.write("재료가 추가되었습니다!")
        _, ingredients = get_data()
        st.session_state.ingredient_df = pd.DataFrame(ingredients, columns=["ID", "IngredientName", "Price"])

    st.header("현재 메뉴 및 재료")

    st.subheader("메뉴")
    st.dataframe(st.session_state.menu_df, use_container_width=True)

    st.subheader("재료")
    if "Quantity" not in st.session_state.ingredient_df.columns:
        st.session_state.ingredient_df["Quantity"] = 0

    if "TotalCost" not in st.session_state.ingredient_df.columns:
        st.session_state.ingredient_df["TotalCost"] = 0

    edited_df = st.data_editor(st.session_state.ingredient_df, num_rows="dynamic")
    if edited_df is not None:
        st.session_state.ingredient_df = edited_df
        # 계산된 총 비용 업데이트
        st.session_state.ingredient_df["TotalCost"] = st.session_state.ingredient_df["Price"] * st.session_state.ingredient_df["Quantity"]

    st.header("검색 기능")

    search_query = st.text_input(
        "",
        "",
        key="search",
        placeholder="검색어를 입력하세요",
        help="Enter a search term to filter the results."
    )

    if search_query:
        filtered_menu_df = st.session_state.menu_df[st.session_state.menu_df.apply(lambda row: search_query.lower() in row.astype(str).str.lower().to_string(), axis=1)]
        filtered_ingredient_df = st.session_state.ingredient_df[st.session_state.ingredient_df.apply(lambda row: search_query.lower() in row.astype(str).str.lower().to_string(), axis=1)]
    else:
        filtered_menu_df = st.session_state.menu_df
        filtered_ingredient_df = st.session_state.ingredient_df

    st.subheader("검색된 메뉴")
    st.dataframe(filtered_menu_df, use_container_width=True)

    st.subheader("검색된 재료")
    st.dataframe(filtered_ingredient_df, use_container_width=True)

    st.header("견적서 계산기")
    st.subheader("재료 목록과 수량 조정")

    total_material_cost = calculate_material_costs(st.session_state.ingredient_df)

    st.subheader("재료 목록")
    st.table(st.session_state.ingredient_df[['IngredientName', 'Price', 'Quantity', 'TotalCost']])

    st.subheader("총 재료 비용")
    st.write(f'총 재료 비용: {total_material_cost} 원')

    num_people = st.number_input('인원 수 입력', min_value=0, step=1)
    total_labor_cost = calculate_labor_cost(num_people)

    st.subheader('인건비')
    st.write(f'인원 수: {num_people} 명')
    st.write(f'인건비: {total_labor_cost} 원')

    total_cost = calculate_total_cost(total_material_cost, total_labor_cost)
    st.write(f'총 비용: {total_cost} 원')

    if st.button('엑셀로 저장'):
        excel_file_path = save_to_excel(st.session_state.ingredient_df, total_material_cost, total_labor_cost, total_cost)
        st.success(f"엑셀 파일이 '{excel_file_path}'로 저장되었습니다.")

def calculate_material_costs(df):
    df['TotalCost'] = df['Price'] * df['Quantity']
    return df['TotalCost'].sum()

def calculate_labor_cost(num_people):
    labor_cost_per_person = 100000
    total_labor_cost = labor_cost_per_person * num_people
    return total_labor_cost

def calculate_total_cost(total_material_cost, total_labor_cost):
    return total_material_cost + total_labor_cost

if __name__ == "__main__":
    main()
