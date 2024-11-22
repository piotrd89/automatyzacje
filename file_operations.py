import os
import pandas as pd
import shutil
from tqdm import tqdm


def excel_to_csv(input_folder: str, output_folder: str) -> None:
    """
    Funkcja konwertująca pliki Excel ze wskazanej lokalizacji do formatu CSV

    Args:
        input_folder (str): ścieżka do plików wejściowych Excel
        output_folder (str): ścieżka dla wygenerowanych plików wyjściowych CSV

    Raises:
        FileExistsError: weryfikacja, czy wskazany plik wejściowy Excel istnieje
    """
    # Sprawdzamy, czy katalog źródłowy istnieje
    if not os.path.exists(input_folder):
        raise FileExistsError(f'Wskazany katalog źródłowy nie istnieje!')
    
    # Sprawdzamy czy katalog docelowy istnieje, jeśli nie, to go tworzymy
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f'Katalog docelowy został utworzony \n{output_folder}')

    # Licznik przekonwertowanych plików
    count_files=0
    excel_files = [file for file in os.listdir(input_folder) if file_name.endswith('.xlsx') or file_name.endswith('.xls')]

    # Przechodzimy przez wszystkie pliki w katalogu wejściowym
    for file_name in os.listdir(input_folder):
        # Sprawdzamy czy pliki mają rozszerzenia *.xls lub *xlsx
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            # Pełna ścieżka do pliku
            input_path = os.path.join(input_folder, file_name)

            # Odczytujemy zawartość pliku
            try:
                # Wczytujemy plik Excel do ramki
                df = pd.read_excel(input_path)

                # Zmieniamy rozszerzenie na *.csv
                output_filename = os.path.splitext(file_name)[0] + '.csv'

                # Ścieżka do pliku wyjściowego
                output_path = os.path.join(output_folder, output_filename)

                # Zapisujemy ramkę jako CSV z kodowaniem UTF-8 oraz indeksowaniem i separatorem ';'
                df.to_csv(output_path, index=True, sep=';', encoding='utf-8')

                print(f'Utworzono plik: {output_filename}')
                count_files += 1
                print(f'{count_files}/{len(excel_files)} ')
            
            except Exception as e:
                print(f'Błąd podczas przetwarzania pliku {file_name}: {e}')
    print(f'Przetworzona liczba plików: {count_files}')

    # Przykład użycia funkcji konwertującej pliki xls, xlsx do formatu csv
    # input_folder = 'ścieżka\\do\\input'
    # output_folder = 'ścieżka\\do\\output'
    # excel_to_csv(input_folder, output_folder)

    # Sprawdzamy przy pomocy modułu Pandas, czy przykładowy plik csv jest poprawny
    # tdf = pd.read_csv('ścieżka\\do\\folderu_z_plikami\plik.csv', delimiter=';')
    # tdf.head()

def csv_to_excel_progressbar(input_folder: str, output_folder: str, delimiter: str, output_extension: str) -> None:
    """
    Funkcja konwertująca pliki CSV do określonego przez użytkownika formatu xlsx lub xls.
    W funkcji wykorzystany jest moduł tqdm do wyświetlania dynamicznie generowanego paska postępu

    Args:
        input_folder (str): ścieżka do folderu wejściowego z plikami do konwersji
        output_folder (str): ścieżka do folderu wyjściowego, gdzie generowane mają być skonwertowane pliki
        delimiter (str): określami separator dla pliku CSV np. ; lub ,
        output_extension (str): format wyjściowy, do któego skonwertowany ma być plik Excel

    Raises:
        FileExistsError: obsługa błędu w przypadku kiedy wskazana ścieżka wejściowa nie istnieje
    """
    # Sprawdzamy, czy katalog źródłowy istnieje
    if not os.path.exists(input_folder):
        raise FileExistsError(f'Wskazany katalog źródłowy nie istnieje!')
    
    # Sprawdzamy czy katalog docelowy istnieje, jeśli nie, to go tworzymy
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f'Katalog docelowy został utworzony \n{output_folder}')

    # Licznik przekonwertowanych plików
    #count_files = 0

    # Pobieramy listę plików CSV z katalogu źródłowego
    csv_files = [file for file in os.listdir(input_folder) if file.endswith('.csv')]

    # Przechodzimy przez wszystkie pliki w katalogu wejściowym
    with tqdm(total=len(csv_files), desc='Konwersja CSV na Excel') as pbar:
        #for file in os.listdir(input_folder):
        #if file.endswith('.csv'):
        for file in csv_files:
            # Ścieżka źródłowa pliku CSV
            csv_file = os.path.join(input_folder, file)
            try:
                # Wczytujemy plik CSV do ramki danych (Data Frame), uwzględniamy separator
                df = pd.read_csv(csv_file, delimiter=delimiter)
                # Ścieżka docelowa pliku Excel
                excel_file = os.path.join(output_folder, os.path.splitext(file)[0] + output_extension)
                # Zapisujemy ramkę danych do pliku Excel. Domyślnie bez indeksu
                df.to_excel(excel_file, index=False)
                # Aktualizujemy pasek postępu
                pbar.update(1)
                #count_files += 1
                #print(f'Przekonwertowano: {csv_file} -> {excel_file}')
            except Exception as e:
                print(f'Błąd podczas przetwarzania pliku {file}: {e}')
    
    #print(f'Przekonwertowanych plików: {count_files}')

# csv_to_excel_progressbar('ścieżka\\do\\input\\pliki_csv', 'ścieżka\\do\\output\\pliki_excel',';','.xlsx')

def csv_to_excel(input_folder: str, output_folder: str, delimiter: str, output_extension: str) -> None:
    """
    Funkcja konwertująca pliki CSV do określonego przez użytkownika formatu *.xlsx lub *.xls

    Args:
        input_folder (str): ścieżka do folderu wejściowego z plikami do konwersji
        output_folder (str): ścieżka do folderu wyjściowego, gdzie generowane mają być skonwertowane pliki
        delimiter (str): określami separator dla pliku CSV np. ; lub ,
        output_extension (str): format wyjściowy, do któego skonwertowany ma być plik Excel

    Raises:
        FileExistsError: obsługa błędu w przypadku kiedy wskazana ścieżka wejściowa nie istnieje
    """
    # Sprawdzamy, czy katalog źródłowy istnieje
    if not os.path.exists(input_folder):
        raise FileExistsError(f'Wskazany katalog źródłowy nie istnieje!')
    
    # Sprawdzamy czy katalog docelowy istnieje, jeśli nie, to go tworzymy
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f'Katalog docelowy został utworzony \n{output_folder}')

    # Licznik przekonwertowanych plików
    count_files = 0

    # Pobieramy listę plików CSV z katalogu źródłowego
    #csv_files = [file for file in os.listdir(input_folder) if file.endswith('.csv')]

    # Przechodzimy przez wszystkie pliki w katalogu wejściowym
    for file in os.listdir(input_folder):
        if file.endswith('.csv'):
            try:
                # Ścieżka źródłowa pliku CSV
                csv_file = os.path.join(input_folder, file)
                # Wczytujemy plik CSV do ramki danych (Data Frame), uwzględniamy separator
                df = pd.read_csv(csv_file, delimiter=delimiter)
                # Ścieżka docelowa pliku Excel
                excel_file = os.path.join(output_folder, os.path.splitext(file)[0] + output_extension)
                # Zapisujemy ramkę danych do pliku Excel. Domyślnie bez indeksu
                df.to_excel(excel_file, index=False)
                count_files += 1
                print(f'Przekonwertowano: {csv_file} -> {excel_file}')
                #print(f'Przekonwertowano: {count_files} z {len(csv_files)}')
            except Exception as e:
                print(f'Błąd podczas przetwarzania pliku {file}: {e}')
                
    print(f'Przekonwertowano plików: {count_files}')
  
# csv_to_excel('ścieżka\\do\\input', 'ścieżka\\do\\output',';','.xlsx')

def download_from_folder(input_path: str, output_folder: str, file_extensions: list) -> None:
    """
    Funkcja kopiująca wszystkie pliki o wybranych rozszerzeniach z całego drzewa katalogów
    do wskazanej przez użytkownika lokalizacji

    Args:
        input_path (str): ścieżka do folderu wejściowego z plikami
        output_folder (str): ścieżka do folderu wyjściowego, gdzie pliki zostaną skopiowane
        file_extensions (list): lista rozszerzeń plików do skopiowania

    Raises:
        FileExistsError: obsługa błędu w przypadku kiedy wskazana ścieżka wejściowa nie istnieje
    """
   # Sprawdzamy, czy katalog źródłowy istnieje
    if not os.path.exists(input_path):
        raise FileExistsError(f'Wskazany katalog źródłowy nie istnieje!')
    
    # Sprawdzamy czy katalog docelowy istnieje, jeśli nie, to go tworzymy
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f'Katalog docelowy został utworzony \n{output_folder}')

    # Licznik skopiowanych elementów
    count_copies = 0

    # Przechodzimy przez wszystkie foldery i pliki w folderze źródłowym
    for root, dirs, files in os.walk(input_path):
        for file in files:

            # Sprawdzamy rozszerzenie pliku
            if file_extensions and not any(file.endswith(ext) for ext in file_extensions):
                continue

            # Ścieżka do pliku źródłowego
            input_file = os.path.join(root, file)
            # Ścieżka docelowa dla pliku
            output_file = os.path.join(output_folder, file)

            # Jeśli plik istnieje w katalogu docelowym to zmieniamy jego nazwę, żeby go nie nadpisać
            base, extension = os.path.splitext(file)
            counter = 1
            while os.path.exists(output_file):
                output_file = os.path.join(output_folder,f'{base}_{counter}{extension}')
                counter += 1
            
            # Kopiowanie pliku
            shutil.copy2(input_file, output_file)
            count_copies += 1
            print(f'Skopiowano: {file} -> {output_folder}')
    print(f'Skopiowanych plików: {count_copies}')

# Przykład użycia funkcji kopiującej wszystkie pliki xls,xlsx z zadanej lokalizacji do wskazanego folderu docelowego
# download_from_folder('ścieżka\\do\\input','ścieżka\\do\\output',['.xls','.xlsx'])