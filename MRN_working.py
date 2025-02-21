import pandas as pd
from rapidfuzz import fuzz
import unicodedata

# Шлях до файлу Excel
file_path = r'C:\Users\Home\Desktop\MRN 07(відновлено автоматично).xlsx'

# Завантаження даних з обох аркушів
search_df = pd.read_excel(file_path, sheet_name='Search', dtype=str, header=0)
data_df = pd.read_excel(file_path, sheet_name='DATA', dtype=str, header=0)

# Видаляємо зайві пробіли у назвах колонок
search_df.columns = search_df.columns.str.strip()
data_df.columns = data_df.columns.str.strip()

# Виводимо список колонок, щоб знати їхній порядок
print("🔍 Колонки у Search:", list(enumerate(search_df.columns)))
print("🔍 Колонки у DATA:", list(enumerate(data_df.columns)))

# Функція для очищення тексту
def clean_text(text):
    if pd.isna(text):
        return ""
    text = str(text).strip().lower()
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(c for c in text if not unicodedata.combining(c))
    return text

# Використання номерів колонок
nr_przesylki_col = 0  # Колонка Nr_przesylki в Search
nr26_lp_col = 0       # Колонка Nr26 (LP) у DATA
opis_towaru_search_col = 7  # Колонка Opis_towaru у Search
opis_towaru_data_col = 3    # Колонка Opis Towaru у DATA
nr_star_col = 11  # Колонка Nr* у Search
mrn_search_col = 12  # Колонка MRN у Search
mrn_data_col = 1  # Колонка MRN у DATA
nr_data_col = 2  # Колонка Nr у DATA

# Очистка тексту у відповідних колонках
search_df.iloc[:, nr_przesylki_col] = search_df.iloc[:, nr_przesylki_col].apply(clean_text)
data_df.iloc[:, nr26_lp_col] = data_df.iloc[:, nr26_lp_col].apply(clean_text)
search_df.iloc[:, opis_towaru_search_col] = search_df.iloc[:, opis_towaru_search_col].apply(clean_text)
data_df.iloc[:, opis_towaru_data_col] = data_df.iloc[:, opis_towaru_data_col].apply(clean_text)

# Замінюємо NaN на порожні значення у Nr* і MRN
search_df.iloc[:, nr_star_col] = search_df.iloc[:, nr_star_col].replace({'#N/D': "", pd.NA: ""}).astype(str)
search_df.iloc[:, mrn_search_col] = search_df.iloc[:, mrn_search_col].replace({'#N/D': "", pd.NA: ""}).astype(str)
data_df.iloc[:, mrn_data_col] = data_df.iloc[:, mrn_data_col].replace({'#N/D': "", pd.NA: ""}).astype(str)

# Проходимо по всіх рядках Search, де Nr* = 0
for index, search_row in search_df.iterrows():
    if str(search_row.iloc[nr_star_col]).strip() not in ["0", "", "nan"]:
        continue

    # Пошук відповідних рядків у DATA за Nr_przesylki та Nr26 (LP)
    matching_rows = data_df[data_df.iloc[:, nr26_lp_col] == search_row.iloc[nr_przesylki_col]]

    # Якщо немає збігів, переходимо до наступного рядка
    if matching_rows.empty:
        continue

    # Фільтрація по MRN (замінюємо NaN на порожній рядок)
    search_mrn = search_row.iloc[mrn_search_col] if pd.notna(search_row.iloc[mrn_search_col]) else ""
    matching_rows.iloc[:, mrn_data_col] = matching_rows.iloc[:, mrn_data_col].fillna("")
    matching_rows = matching_rows[matching_rows.iloc[:, mrn_data_col] == search_mrn]

    # Якщо немає збігів, переходимо до наступного рядка
    if matching_rows.empty:
        continue

    # Перевірка часткової схожості опису
    for _, match_row in matching_rows.iterrows():
        similarity = fuzz.partial_ratio(search_row.iloc[opis_towaru_search_col], match_row.iloc[opis_towaru_data_col])
        
        # Якщо схожість більше 70%, оновлюємо значення в Nr*
        if similarity > 70:
            search_df.at[index, search_df.columns[nr_star_col]] = str(match_row.iloc[nr_data_col]).strip() if pd.notna(match_row.iloc[nr_data_col]) else "0"

# Збереження оновленого файлу
output_path = r'C:\Users\Home\Desktop\newone.xlsx'
search_df.to_excel(output_path, index=False)

print(f"✅ Дані успішно оновлені! Оновлений файл збережено: {output_path}")
