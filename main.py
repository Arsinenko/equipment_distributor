import pandas as pd 
import json
import re


regex_rules = {
    'foreign_lang': r'иностран',
    'biology_safety': r'биолог|обзр|безопасн', 
    'physics': r'физик',                       
    'chemistry': r'химия|химии',               
    'informatics': r'информат',                
    'it_material': r'мастерск|материал',       
    'universal': r'универсал|класс|математ|русск|литерат|истор|географ|обществозн|начальн'
}

def get_subject_key(text):
    """Определяет ключ комнаты на основе регулярного выражения"""
    if not isinstance(text, str):
        return None
    
    for room_key, pattern in regex_rules.items():
        if re.search(pattern, text, re.IGNORECASE):
            return room_key      
    return None

def load_equipment_db(filepath: str) -> pd.DataFrame:
    df = pd.read_excel(filepath)
    df = df.rename(columns={
        "№": "Number",
        "Раздел": "Section",
        "Помещение": "Room",
        "Наименование_2025": "Name",
        "Ед. изм.": "Unit",
        "КОЛ-ВО на кабинет": "Count",
    })
    df.columns = df.columns.str.strip()
    return df 

def get_school_models(filepath: str) -> dict:
    with open(filepath, 'r', encoding='utf-8') as f:
        data = json.load(f)
    return data

def choose_model(school_models: dict, model: str) -> dict:
    return school_models[model] 

def calculate_needed_equipment(df: pd.DataFrame, rooms: dict) -> pd.DataFrame:
    # 1. Определяем категорию (subject_room) для каждой строки
    df['Subject_room'] = df['Room'].apply(get_subject_key)
    
    # 2. Получаем множитель, сопоставляя найденный subject_room со словарем rooms из JSON
    # Если комната найдена, берем значение, если нет — 0
    df['Multiplier'] = df['Subject_room'].map(rooms).fillna(0)
    
    # 3. Считаем итоговое количество
    df['Total_Count'] = df['Count'] * df['Multiplier']
    
    # 4. Формируем итоговый датафрейм, добавляя subject_room в выборку
    final_df = df[['Name', 'Total_Count', 'Unit', 'Subject_room']].rename(columns={'Total_Count': 'Count'})
    
    # 5. Фильтруем нулевые значения
    final_df = final_df[final_df['Count'] > 0]
    
    # 6. Группируем. Важно добавить 'subject_room' в группировку, 
    # иначе pandas не будет знать, какую категорию оставить при суммировании
    final_df = final_df.groupby(['Name', 'Unit', 'Subject_room'], as_index=False)['Count'].sum()
    
    return final_df

def main():
    model_name = '7-11'
    school_name_path = 'care.json'
    school_models: dict = get_school_models(school_name_path)
    model_data: dict = choose_model(school_models, model_name)
    rooms: dict = model_data['rooms']
    df = load_equipment_db('equipment.xlsx')
    final_df = calculate_needed_equipment(df, rooms)
    print(final_df)
    
    
    
if __name__ == '__main__':
    main()
    
    
    

    