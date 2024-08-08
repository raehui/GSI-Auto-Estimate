import sqlite3
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font

# SQLite 데이터베이스와 연결하는 함수
def connect_db():
    return sqlite3.connect('restaurant_menu.db')

# SQLite 데이터베이스 초기화 함수
def reset_database():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    # 기존 테이블 삭제 (데이터베이스를 초기화하기 위해)
    c.execute('DROP TABLE IF EXISTS MenuIngredients')
    c.execute('DROP TABLE IF EXISTS Menus')
    c.execute('DROP TABLE IF EXISTS Ingredients')

    conn.commit()
    conn.close()

    # 데이터베이스 다시 생성
    create_database()
    insert_data()

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
    try:
        c.execute('SELECT * FROM Menus')
        menus = c.fetchall()
        c.execute('SELECT * FROM Ingredients')
        ingredients = c.fetchall()
        c.execute('''SELECT mi.MenuID, m.MenuName, mi.IngredientID, i.IngredientName, mi.Quantity, i.Price
                     FROM MenuIngredients mi
                     JOIN Menus m ON mi.MenuID = m.MenuID
                     JOIN Ingredients i ON mi.IngredientID = i.IngredientID''')
        menu_ingredients = c.fetchall()
        conn.close()
        return menus, ingredients, menu_ingredients
    except sqlite3.Error as e:
        st.error(f"SQLite error: {e}")
        return [], [], []  # 에러 발생 시 빈 리스트 반환

def create_database():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    # Menus 테이블 생성
    c.execute('''CREATE TABLE IF NOT EXISTS Menus (
                    MenuID INTEGER PRIMARY KEY AUTOINCREMENT,
                    MenuName TEXT NOT NULL,
                    DateSubmitted TEXT NOT NULL
                )''')

    # Ingredients 테이블 생성
    c.execute('''CREATE TABLE IF NOT EXISTS Ingredients (
                    IngredientID INTEGER PRIMARY KEY AUTOINCREMENT,
                    IngredientName TEXT NOT NULL,
                    Price REAL NOT NULL
                )''')

    # MenuIngredients 테이블 생성
    c.execute('''CREATE TABLE IF NOT EXISTS MenuIngredients (
                    MenuID INTEGER,
                    IngredientID INTEGER,
                    Quantity REAL NOT NULL,
                    FOREIGN KEY (MenuID) REFERENCES Menus(MenuID),
                    FOREIGN KEY (IngredientID) REFERENCES Ingredients(IngredientID)
                )''')

    conn.commit()
    conn.close()

def insert_data():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    # 메뉴 데이터 추가
    menus = [
        ('김치찌개', '2024-01-01'),
        ('된장찌개', '2024-01-01'),
        ('순두부찌개', '2024-01-01')
    ]
    c.executemany('INSERT INTO Menus (MenuName, DateSubmitted) VALUES (?, ?)', menus)

    # 재료 데이터 추가
    ingredients = [
        ('김치', 2.0),
        ('양파', 0.5),
        ('대파', 0.3),
        ('마늘', 0.2),
        ('고춧가루', 0.1),
        ('멸치', 1.0),
        ('된장', 0.7),
        ('소금', 0.05),
        ('돼지고기', 3.0),
        ('두부', 1.5),
        ('물', 0.0),
        ('김칫국물', 0.0),
        ('청양고추', 0.1),
        ('감자', 0.8),
        ('고추', 0.2),
        ('버섯', 1.0),
        ('호박', 0.6),
        ('순두부', 2.5),
        ('계란', 0.3),
    ]
    c.executemany('INSERT INTO Ingredients (IngredientName, Price) VALUES (?, ?)', ingredients)

    # 메뉴-재료 관계 데이터 추가
    menu_ingredients = [
        # 김치찌개 재료
        (1, 1, 1.0),   # 김치찌개, 김치 1.0 단위
        (1, 2, 1.0),   # 김치찌개, 양파 0.5 단위
        # 된장찌개 재료들
        (2, 7, 1.0),   # 된장찌개, 된장 1.0 단위
        (2, 14, 1.0),  # 된장찌개, 감자 1.0 단위
        # 순두부찌개 재료들
        (3, 18, 1.0),  # 순두부찌개, 순두부 1.0 단위
        (3, 2, 1.0),   # 순두부찌개, 양파 0.5 단위
    ]
    c.executemany('INSERT INTO MenuIngredients (MenuID, IngredientID, Quantity) VALUES (?, ?, ?)', menu_ingredients)

    conn.commit()
    conn.close()


# 메뉴 총 비용 계산 함수
def calculate_menu_costs(menu_df, menu_ingredients_df, ingredient_df):
    menu_costs = []
    for _, menu_row in menu_df.iterrows():
        menu_id = menu_row["ID"]
        quantity = menu_row["Quantity"]
        total_cost = 0

        # 메뉴에 포함된 재료를 찾아서 총 비용 계산
        menu_ingredients = menu_ingredients_df[menu_ingredients_df["MenuID"] == menu_id]
        for _, mi_row in menu_ingredients.iterrows():
            ingredient_id = mi_row["IngredientID"]
            ingredient_price = ingredient_df[ingredient_df["ID"] == ingredient_id]["Price"].values[0]
            ingredient_quantity = mi_row["Quantity"] * quantity
            total_cost += ingredient_price * ingredient_quantity

        menu_costs.append(total_cost)
    
    return menu_costs

def main():
    st.set_page_config(page_title="레스토랑 메뉴 관리")
    st.title("레스토랑 메뉴 및 재료 관리 시스템")

    if st.sidebar.button("초기화"):
        reset_database()
        st.sidebar.success("데이터베이스가 초기화되었습니다!")

    if "menu_df" not in st.session_state or "ingredient_df" not in st.session_state:
        menus, ingredients, menu_ingredients = get_data()

        # 필요 시 필요한 열만 선택하여 DataFrame 생성
        if menus and len(menus[0]) == 3:
            st.session_state.menu_df = pd.DataFrame(menus, columns=["ID", "MenuName", "DateSubmitted"]).drop(columns=["DateSubmitted"])
        else:
            st.session_state.menu_df = pd.DataFrame(menus, columns=["ID", "MenuName"])

        st.session_state.ingredient_df = pd.DataFrame(ingredients, columns=["ID", "IngredientName", "Price"])
        st.session_state.menu_ingredients_df = pd.DataFrame(menu_ingredients, columns=["MenuID", "MenuName", "IngredientID", "IngredientName", "Quantity", "Price"])

        # 메뉴별 수량 및 총 비용 초기화
        st.session_state.menu_df["Quantity"] = 0
        st.session_state.menu_df["TotalCost"] = 0

    st.sidebar.header("메뉴 추가")
    with st.sidebar.form("메뉴 추가 양식"):
        menu_name = st.text_area("메뉴 이름")
        menu_submitted = st.form_submit_button("메뉴 추가")

    if menu_submitted:
        insert_menu(menu_name)
        st.sidebar.write("메뉴가 추가되었습니다!")
        menus, _, _ = get_data()
        if menus and len(menus[0]) == 3:
            st.session_state.menu_df = pd.DataFrame(menus, columns=["ID", "MenuName", "DateSubmitted"]).drop(columns=["DateSubmitted"])
        else:
            st.session_state.menu_df = pd.DataFrame(menus, columns=["ID", "MenuName"])
        st.session_state.menu_df["Quantity"] = 0
        st.session_state.menu_df["TotalCost"] = 0

    st.sidebar.header("재료 추가")
    with st.sidebar.form("재료 추가 양식"):
        ingredient_name = st.text_area("재료 이름")
        price = st.number_input("가격", min_value=0.0, step=0.1)
        ingredient_submitted = st.form_submit_button("재료 추가")

    if ingredient_submitted:
        insert_ingredient(ingredient_name, price)
        st.sidebar.write("재료가 추가되었습니다!")
        _, ingredients, _ = get_data()
        st.session_state.ingredient_df = pd.DataFrame(ingredients, columns=["ID", "IngredientName", "Price"])

    st.header("현재 메뉴 및 재료")
    
    # # 기존 메뉴, 업데이트된 메뉴, 재료 테이블 삭제
    # st.subheader("메뉴")
    # menu_df = st.session_state.menu_df
    # edited_menu_df = st.data_editor(menu_df, num_rows="dynamic")
    # if edited_menu_df is not None:
    #     st.session_state.menu_df = edited_menu_df

    # # 메뉴 총 비용 계산
    # menu_costs = calculate_menu_costs(st.session_state.menu_df, st.session_state.menu_ingredients_df, st.session_state.ingredient_df)
    # st.session_state.menu_df["TotalCost"] = menu_costs

    # # 메뉴 데이터프레임을 표시
    # st.subheader("업데이트된 메뉴")
    # st.dataframe(st.session_state.menu_df, use_container_width=True)

    # st.subheader("재료")
    # if "Quantity" not in st.session_state.ingredient_df.columns:
    #     st.session_state.ingredient_df["Quantity"] = 0

    # if "TotalCost" not in st.session_state.ingredient_df.columns:
    #     st.session_state.ingredient_df["TotalCost"] = 0

    # edited_ingredient_df = st.data_editor(st.session_state.ingredient_df, num_rows="dynamic")
    # if edited_ingredient_df is not None:
    #     st.session_state.ingredient_df = edited_ingredient_df

    # 메뉴 선택 및 수량 입력
    st.subheader("메뉴 선택 및 수량 입력")
    selected_menu = st.multiselect("메뉴 선택", st.session_state.menu_df["MenuName"].tolist())
    st.write("선택한 메뉴:")
    for menu_name in selected_menu:
        menu_id = st.session_state.menu_df[st.session_state.menu_df["MenuName"] == menu_name]["ID"].values[0]
        quantity = st.number_input(f"{menu_name} 수량", min_value=0, step=1, key=menu_id)
        st.session_state.menu_df.loc[st.session_state.menu_df["ID"] == menu_id, "Quantity"] = quantity

    # 선택한 메뉴의 재료를 조회하고 업데이트된 재료 테이블에 표시
    filtered_ingredients = st.session_state.menu_ingredients_df[st.session_state.menu_ingredients_df["MenuName"].isin(selected_menu)]
    filtered_ingredient_ids = filtered_ingredients["IngredientID"].unique()
    updated_ingredients = st.session_state.ingredient_df[st.session_state.ingredient_df["ID"].isin(filtered_ingredient_ids)].copy()
    updated_ingredients["Include"] = False  # 체크박스 열 추가

    # Quantity는 수정 가능하게 설정
    edited_updated_ingredients = st.data_editor(
        updated_ingredients,
        num_rows="dynamic",
        column_config={
            "Quantity": st.column_config.NumberColumn("수량"),
            "Include": st.column_config.CheckboxColumn("Include")
        }
    )

    if edited_updated_ingredients is not None:
        st.session_state.ingredient_df.update(edited_updated_ingredients)

    st.subheader("업데이트된 재료")
    st.dataframe(st.session_state.ingredient_df, use_container_width=True)
    
    # 체크박스가 체크된 재료를 메뉴 선택 및 수량 입력 테이블에 추가
    selected_ingredients = edited_updated_ingredients[edited_updated_ingredients["Include"] == True]
    st.subheader("선택된 재료")
    st.dataframe(selected_ingredients, use_container_width=True)
    
    st.session_state.ingredient_df["Quantity"] = 0
    st.session_state.ingredient_df["TotalCost"] = 0

    for _, row in st.session_state.menu_df.iterrows():
        menu_id = row["ID"]
        quantity = row["Quantity"]
        menu_ingredients_df = st.session_state.menu_ingredients_df[st.session_state.menu_ingredients_df["MenuID"] == menu_id]
        for _, mi_row in menu_ingredients_df.iterrows():
            ingredient_id = mi_row["IngredientID"]
            unit_quantity = mi_row["Quantity"] * quantity

            if ingredient_id in st.session_state.ingredient_df["ID"].values:
                ingredient_idx = st.session_state.ingredient_df[st.session_state.ingredient_df["ID"] == ingredient_id].index[0]
                st.session_state.ingredient_df.at[ingredient_idx, "Quantity"] += unit_quantity
                st.session_state.ingredient_df.at[ingredient_idx, "TotalCost"] = st.session_state.ingredient_df.at[ingredient_idx, "Quantity"] * st.session_state.ingredient_df.at[ingredient_idx, "Price"]


    st.header("견적서 계산기")

    def calculate_material_costs(df):
        return df["TotalCost"].sum() if "TotalCost" in df.columns else 0

    def calculate_labor_cost(num_people):
        labor_cost_per_person = 10000
        return num_people * labor_cost_per_person

    total_material_cost = calculate_material_costs(st.session_state.ingredient_df)
    st.write(f'총 재료 비용: {total_material_cost:,} 원')

    num_people = st.number_input('인원 수 입력', min_value=0, step=1)
    total_labor_cost = calculate_labor_cost(num_people)
    st.write(f'인건비: {total_labor_cost:,} 원')

    total_cost = total_material_cost + total_labor_cost
    st.write(f'총 비용: {total_cost:,} 원')

    if st.button('엑셀로 저장'):
        excel_file_path = save_to_excel(st.session_state.ingredient_df, total_material_cost, total_labor_cost, total_cost)
        st.success(f"엑셀 파일이 '{excel_file_path}'로 저장되었습니다.")

def save_to_excel(materials_df, total_material_cost, total_labor_cost, total_cost):
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "표제"
    # 스타일 설정
    thin = Side(border_style="thin", color="000000")
    header_font = Font(bold=True)
    center_align = Alignment(horizontal="center")
    left_align = Alignment(horizontal="left")

    # 표제 부분
    ws1['A1'] = "견적서"
    ws1['A1'].font = Font(size=16, bold=True)
    ws1['A1'].alignment = Alignment(horizontal="center")

    ws1['A7'] = "합계금액:"
    ws1['B7'] = f"{total_cost:,} 원"

    # 메뉴 및 재료 목록 추가
    menus_df = pd.DataFrame(get_data()[0], columns=["ID", "MenuName", "DateSubmitted"]).drop(columns=["DateSubmitted"])
    if not menus_df.empty:
        ws1.append(["메뉴"])
        for idx, row in menus_df.iterrows():
            ws1.append([row['MenuName']])

    ws1.append(['번호', '재료명', '수량', '단위', '단가 (KRW)', '합계 (KRW)'])

    for idx, row in materials_df.iterrows():
        ws1.append([row['ID'], row['IngredientName'], row['Quantity'], '단위', row['Price'], row['TotalCost']])

    # 합계 및 인건비 추가
    ws1.append(['', '', '', '', '재료 총 비용', f"{total_material_cost:,} 원"])
    ws1.append(['', '', '', '', '인건비', f"{total_labor_cost:,} 원"])
    ws1.append(['', '', '', '', '전체 총 비용', f"{total_cost:,} 원"])

    excel_file_path = '견적서.xlsx'
    wb.save(excel_file_path)
    return excel_file_path

if __name__ == "__main__":
    main()