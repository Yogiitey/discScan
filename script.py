import os
from xml.parsers.expat import errors
import pandas as pd
from datetime import datetime
import time
import sys
from pymediainfo import MediaInfo



# Logika pozwalająca odnaleźć DLL wewnątrz pliku .exe
if getattr(sys, 'frozen', False):
    # Jeśli program jest spakowany przez PyInstallera, pliki są w folderze tymczasowym _MEIPASS[cite: 1]
    base_path = sys._MEIPASS
else:
    # Jeśli uruchamiasz skrypt normalnie przez 'python script.py'[cite: 1]
    base_path = os.path.dirname(os.path.abspath(__file__))

DLL_PATH = os.path.join(base_path, "MediaInfo.dll")

def get_file_info():
    lista_danych = []
    index = 0

    if input("Czy chcesz podać konkretną ścieżkę do katalogu? (y/n): ").lower() == "y":
        katalog = input("Podaj ścieżkę do katalogu: ")
    else:
        litera = input("Podaj literę dysku (np. D): ").upper()
        katalog = f"{litera}:\\"

    print(f"Przeszukuję katalog: {katalog}, to może chwilę potrwać...")
    start_time = datetime.now()

    #lista rozszerzeń, dla których będzie uruchamiana funkcja get_metadata(). Dla innych plików zostaną wstawione puste wartości. Dzięki temu skrypt będzie działał szybciej, ponieważ nie będzie próbował pobierać metadanych dla każdego pliku, tylko dla tych, które są potencjalnie multimedialne.
    video_exts = ('.mp4', '.mkv', '.avi', '.mov', '.flv', '.wmv')
    image_exts = ('.jpg', '.png', '.jpeg', '.bmp', '.gif')
    audio_exts = ('.mp3', '.wav', '.flac')
    exts_list = video_exts + image_exts + audio_exts

    start_time_speed = time.time()
    last_count = 0
    pliki_na_sekunde = 0

    for root, dirs, files in os.walk(katalog):
        for nazwa_pliku in files:
            # Tworzymy pełną ścieżkę do pliku
            pelna_sciezka = os.path.join(root, nazwa_pliku)
            index += 1
            rozmiar = round((os.path.getsize(pelna_sciezka) / (1024 * 1024)), 1) # Rozmiar w MB, zaokrąglony do 1 miejsca po przecinku

            # Pobieramy datę utworzenia (timestamp)
            timestamp_utworzenia = os.path.getctime(pelna_sciezka)
            
            # Konwertujemy timestamp na czytelny format (np. RRRR-MM-DD HH:MM:SS)
            data_formatowalna = datetime.fromtimestamp(timestamp_utworzenia).strftime('%d.%m.%Y')
            godzina_formatowalna = datetime.fromtimestamp(timestamp_utworzenia).strftime('%H:%M:%S')

            # --- Logika aktualizacji licznika w jednej linii ---
            current_time = time.time()
            elapsed_chunk = current_time - start_time_speed
            
            # Aktualizujemy statystyki co 1 sekundy, żeby licznik nie "migał" zbyt szybko
            if elapsed_chunk >= 1:
                pliki_na_sekunde = int((index - last_count) / elapsed_chunk)
                last_count = index
                start_time_speed = current_time

            # \r sprawia, że kursor wraca na początek linii, end="" zapobiega nowej linii
            sys.stdout.write(f"\r>>> Przetwarzam {pliki_na_sekunde} plik/s | Mam już za sobą: {index} plików")
            sys.stdout.flush()

    


            ##if index % 10000 == 0:
            ##    print(f"Przetworzono {index} plików...")


            
            # Logika "inteligentnego" skanowania
            ext = nazwa_pliku.lower()
            
            if ext.endswith(exts_list):
            # TYLKO TUTAJ uruchamiamy ciężką funkcję pobierz_metadane()
                meta = get_metadata(pelna_sciezka)
            else:
            # Dla plików .txt, .pdf, .dll itp. wstawiamy puste wartości
                meta = {'Szerokość': None, 'Wysokość': None, 'Bitrate (kbps)': None}

            

           
       # Dodajemy słownik z informacjami do listy
            lista_danych.append({
                'index': index,
                'Nazwa pliku': nazwa_pliku,
                'Folder': root,
                'Pełna ścieżka': pelna_sciezka,
                'Rozmiar (MB)': rozmiar,
                'Godzina utworzenia': godzina_formatowalna,
                'Data utworzenia': data_formatowalna,

                
                #kolumna szerokości, wysokości i bitrate
                'Szerokość (px)': meta.get('Szerokość'),
                'Wysokość (px)': meta.get('Wysokość'),
                'Bitrate (kbps)': meta.get('Bitrate (kbps)')
                })
            
    # Po zakończeniu pętli dodajemy nową linię, żeby kolejne komunikaty nie nadpisały licznika
    print("\n--- Skanowanie zakończone! ---")

    # Zamiana listy słowników na DataFrame
    df = pd.DataFrame(lista_danych)

    # Zapis do pliku Excel
    print("Zapisuję dane do twojego excela...")
    df.to_excel('file_info.xlsx', index=False)

    duration = datetime.now() - start_time


    print(f"Zakończono przetwarzanie. Przetworzono {index} plików. Dane zapisano w 'file_info.xlsx'. Czas trwania: {duration}")
    if index >= 10000:
        speed = round(index / duration.total_seconds(), 2)
        print(f"Średnia prędkość skanowania: {speed} plików/s")

        """
    Zapisuje podsumowanie skanowania do pliku tekstowego.
    """
        nazwa_logu = f"raport_skanowania_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
        with open(nazwa_logu, "w", encoding="utf-8") as f:
            f.write("=== RAPORT Z PRZEBIEGU SKANOWANIA ===\n")
            f.write(f"Data raportu: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
            f.write("-" * 40 + "\n")
            f.write(f"Scieżka skanowania: {katalog}\n")
            f.write(f"Czas rozpoczęcia: {start_time.strftime('%H:%M:%S')}\n")
            f.write(f"Czas zakończenia: {datetime.now().strftime('%H:%M:%S')}\n")
            f.write(f"Całkowity czas: {duration}\n")
            f.write("-" * 40 + "\n")
            f.write(f"Przetworzonych plików: {index}\n")
            f.write(f"Średnia prędkość: {speed} plików/s\n")
            f.write("-" * 40 + "\n")
            f.write("Status: Zakończono pomyślnie.\n")

        print(f"\nRaport techniczny zapisano w: {nazwa_logu}")
    

    return True;

def get_metadata(path):
    # Domyślne wartości
    dane = {'Szerokość': None, 'Wysokość': None, 'Bitrate (kbps)': None}

    
    try:
        mi = MediaInfo.parse(path, library_file=DLL_PATH)

        for track in mi.tracks:
            # 1. Obsługa WIDEO - pobieramy tylko wymiary
            if track.track_type == "Video":
                dane['Szerokość'] = track.width
                dane['Wysokość'] = track.height
            
            # 2. Obsługa OBRAZÓW - tylko wymiary
            elif track.track_type == "Image":
                dane['Szerokość'] = track.width
                dane['Wysokość'] = track.height
            
            # 3. Obsługa AUDIO - to tutaj wyciągamy konkretny bitrate ścieżki dźwiękowej
            elif track.track_type == "Audio":
                # Sprawdzamy bitrate konkretnie dla ścieżki audio
                
                br = track.bit_rate or track.nominal_bit_rate
                if br:
                    dane['Bitrate (kbps)'] = int(br) // 1000
                # Niektóre formaty (np. FLAC) używają pola 'sampling_rate' i 'bit_depth', 
                # ale MediaInfo zazwyczaj podaje wyliczony 'bit_rate' dla większości kontenerów.
    except Exception as e:
        print(f"Błąd metadanych dla {path}: {e}")

        #zapisujemy błędy do listy, żeby później móc je zapisać do txt
        with open("errors.txt", "w", encoding="utf-8") as f:
            f.write(f"{path}: {e}\n")
        pass # W razie błędu zwracamy puste wartości
    return dane





def exit_countdown(seconds):
    revolver = 0
    for i in range(seconds, 0, -1):

        if revolver == 0:
            symbol = "/"
            revolver = 1
        elif revolver == 1:
            symbol = "-"
            revolver = 2
        elif revolver == 2:
            symbol = "\\"
            revolver = 3
        elif revolver == 3:
            symbol = "|"
            revolver = 0


        
        
        sys.stdout.write(f"\rZamykanie programu za {i} sekund...  {symbol}")
        sys.stdout.flush()
        time.sleep(1)
    print("\nProgram został zamknięty.")




get_file_info()
exit_countdown(15)