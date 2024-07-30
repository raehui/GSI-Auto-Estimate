import sqlite3

def create_database():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    # Menus 테이블 생성
    c.execute('''CREATE TABLE IF NOT EXISTS Menus (
                    MenuID INTEGER PRIMARY KEY AUTOINCREMENT,
                    MenuName TEXT NOT NULL
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

create_database()

def insert_data():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    # 메뉴 데이터 추가
    menus = [
        ('김치찌개',),
        ('된장찌개',),
        ('순두부찌개',)
    ]
    c.executemany('INSERT INTO Menus (MenuName) VALUES (?)', menus)

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
        (1, 2, 0.5),   # 김치찌개, 양파 0.5 단위
        # 된장찌개 재료들
        (2, 7, 1.0),   # 된장찌개, 된장 1.0 단위
        (2, 14, 1.0),  # 된장찌개, 감자 1.0 단위
        # 순두부찌개 재료들
        (3, 18, 1.0),  # 순두부찌개, 순두부 1.0 단위
        (3, 2, 0.5),   # 순두부찌개, 양파 0.5 단위
    ]
    c.executemany('INSERT INTO MenuIngredients (MenuID, IngredientID, Quantity) VALUES (?, ?, ?)', menu_ingredients)

    conn.commit()
    conn.close()

insert_data()

def view_data():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    #Menus 테이블 데이터 조회
    c.execute('SELECT * FROM Menus')
    menus = c.fetchall()
    print("Menus:")
    for menu in menus:
        print(menu)

    #Ingredients 테이블 데이터 조회
    c.execute('SELECT * FROM Ingredients')
    ingredients = c.fetchall()
    print("\nIngredients:")
    for ingredient in ingredients:
        print(ingredient)

    conn.close()

import sqlite3
import pandas as pd
import streamlit as st

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
    c.execute('''SELECT mi.MenuID, m.MenuName, mi.IngredientID, i.IngredientName, mi.Quantity, i.Price
                 FROM MenuIngredients mi
                 JOIN Menus m ON mi.MenuID = m.MenuID
                 JOIN Ingredients i ON mi.IngredientID = i.IngredientID''')
    menu_ingredients = c.fetchall()
    conn.close()
    return menus, ingredients, menu_ingredients

# 데이터베이스 초기화 및 데이터 삽입 함수
def create_database():
    conn = sqlite3.connect('restaurant_menu.db')
    c = conn.cursor()

    # Menus 테이블 생성
    c.execute('''CREATE TABLE IF NOT EXISTS Menus (
                    MenuID INTEGER PRIMARY KEY AUTOINCREMENT,
                    MenuName TEXT NOT NULL
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
        ('김치찌개',),
        ('된장찌개',),
        ('순두부찌개',)
    ]
    c.executemany('INSERT INTO Menus (MenuName) VALUES (?)', menus)

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
        (1, 2, 0.5),   # 김치찌개, 양파 0.5 단위
        # 된장찌개 재료들
        (2, 7, 1.0),   # 된장찌개, 된장 1.0 단위
        (2, 14, 1.0),  # 된장찌개, 감자 1.0 단위
        # 순두부찌개 재료들
        (3, 18, 1.0),  # 순두부찌개, 순두부 1.0 단위
        (3, 2, 0.5),   # 순두부찌개, 양파 0.5 단위
    ]
    c.executemany('INSERT INTO MenuIngredients (MenuID, IngredientID, Quantity) VALUES (?, ?, ?)', menu_ingredients)

    conn.commit()
    conn.close()

create_database()
insert_data()

def main():
    st.set_page_config(page_title="레스토랑 메뉴 관리")
    st.title("레스토랑 메뉴 및 재료 관리 시스템")

    if "menu_df" not in st.session_state or "ingredient_df" not in st.session_state or "menu_ingredient_df" not in st.session_state:
        menus, ingredients, menu_ingredients = get_data()
        st.session_state.menu_df = pd.DataFrame(menus, columns=["MenuID", "MenuName"])
        st.session_state.ingredient_df = pd.DataFrame(ingredients, columns=["IngredientID", "IngredientName", "Price"])
        st.session_state.menu_ingredient_df = pd.DataFrame(menu_ingredients, columns=["MenuID", "MenuName", "IngredientID", "IngredientName", "Quantity", "Price"])

    st.sidebar.header("메뉴 추가")
    with st.sidebar.form("메뉴 추가 양식"):
        menu_name = st.text_area("메뉴 이름")
        menu_submitted = st.form_submit_button("메뉴 추가")

    if menu_submitted:
        insert_menu(menu_name)
        st.sidebar.write("메뉴가 추가되었습니다!")
        menus, _, _ = get_data()
        st.session_state.menu_df = pd.DataFrame(menus, columns=["MenuID", "MenuName"])

    st.sidebar.header("재료 추가")
    with st.sidebar.form("재료 추가 양식"):
        ingredient_name = st.text_area("재료 이름")
        price = st.number_input("가격", min_value=0.0, step=0.1)
        ingredient_submitted = st.form_submit_button("재료 추가")

    if ingredient_submitted:
        insert_ingredient(ingredient_name, price)
        st.sidebar.write("재료가 추가되었습니다!")
        _, ingredients, _ = get_data()
        st.session_state.ingredient_df = pd.DataFrame(ingredients, columns=["IngredientID", "IngredientName", "Price"])

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

    if st.button('견적서 Excel 파일 생성'):
        create_estimate_excel(st.session_state.menu_ingredient_df, total_material_cost, num_people, total_labor_cost, total_cost)

def calculate_material_costs(df):
    df['TotalCost'] = df['Price'] * df['Quantity']
    return df['TotalCost'].sum()

def calculate_labor_cost(num_people):
    labor_cost_per_person = 100000
    total_labor_cost = labor_cost_per_person * num_people
    return total_labor_cost

def calculate_total_cost(total_material_cost, total_labor_cost):
    return total_material_cost + total_labor_cost

def create_estimate_excel(menu_ingredient_df, total_material_cost, num_people, total_labor_cost, total_cost):
    template_file_path = 'Sample.xlsx'
    xls = pd.ExcelFile(template_file_path)
    df_title = pd.read_excel(xls, sheet_name='표제 ')
    df_details = pd.read_excel(xls, sheet_name='세부')

    # 세부 시트 업데이트
    for i, row in menu_ingredient_df.iterrows():
        df_details.at[i + 3, 'Unnamed: 5'] = row['Quantity']  # 재료 수량
        df_details.at[i + 3, 'Unnamed: 6'] = row['Price']     # 재료 가격
        df_details.at[i + 3, 'Unnamed: 7'] = row['Price'] * row['Quantity'] # 재료 총 비용

    # 엑셀 파일로 저장
    excel_file_path = '견적서.xlsx'
    with pd.ExcelWriter(excel_file_path) as writer:
        df_title.to_excel(writer, sheet_name='표제 ')
        df_details.to_excel(writer, sheet_name='세부')

    st.success('견적서 Excel 파일이 생성되었습니다.')
    st.write(f'[견적서 다운로드]({excel_file_path})')

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import sqlite3

# 페이지 설정을 가장 먼저 설정해야 함
# st.set_page_config(page_title="레스토랑 메뉴 관리")

# SQLite 데이터베이스 연결 함수
def connect_db():
    return sqlite3.connect('restaurant_menu.db')

# 데이터베이스에서 메뉴와 재료 데이터 가져오기
def get_data():
    conn = connect_db()
    c = conn.cursor()
    c.execute('SELECT i.IngredientName, i.Price, mi.Quantity, (i.Price * mi.Quantity) AS TotalCost FROM MenuIngredients mi JOIN Ingredients i ON mi.IngredientID = i.IngredientID')
    data = c.fetchall()
    conn.close()
    return data

# 스트림릿 메인 앱 함수
def main():
    st.title("레스토랑 메뉴 및 재료 관리 시스템")

    data = get_data()
    if data:
        df = pd.DataFrame(data, columns=['IngredientName', 'Price', 'Quantity', 'TotalCost'])

        # 견적서 테이블 형식으로 변환
        df_estimate = pd.DataFrame({
            'Description': df['IngredientName'],
            'Specification': [''] * len(df),  # 빈 값 또는 필요한 정보를 추가할 수 있습니다.
            'Q\'ty': df['Quantity'],
            'Unit': ['Set'] * len(df),  # 또는 다른 단위를 사용할 수 있습니다.
            'Unit Price': df['Price'],
            'Price': df['TotalCost'],
            'Remark': [''] * len(df)  # 빈 값 또는 필요한 정보를 추가할 수 있습니다.
        })

        st.subheader("재료 견적서")
        st.dataframe(df_estimate, use_container_width=True)


        # Q'ty와 Price의 총합 계산
        total_qty = df_estimate['Q\'ty'].sum()
        total_price = df_estimate['Price'].sum()

        st.subheader("소계")
        st.write(f"Q'ty 총합: {total_qty}")
        st.write(f"Price 총합: {total_price} 원")
        
if __name__ == "__main__":
    main()
