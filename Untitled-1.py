import pandas as pd
from datetime import datetime, timedelta

def generate_schedule(excel_path, output_path, year):
    # Excel'den saat verilerini oku
    df_saatler = pd.read_excel(excel_path)
    
    # Hat No ve Ay değerlerini otomatik olarak dosyadan al
    unique_hat_no = df_saatler["Hat No"].unique()
    unique_ay = df_saatler["Ay"].unique()
    
    # Eğer dosyada birden fazla 'Hat No' veya 'Ay' değeri varsa, kullanıcıyı uyar veya işlemi durdur
    if len(unique_hat_no) != 1 or len(unique_ay) != 1:
        raise ValueError("Excel dosyası birden fazla 'Hat No' veya 'Ay' içeriyor. Lütfen dosyayı kontrol edin.")
    
    hat_no = unique_hat_no[0]
    month = unique_ay[0]
    
    # Seçilen ayın başlangıç ve bitiş tarihlerini hesapla
    start_date = datetime(year, month, 1)
    if month == 12:
        end_date = datetime(year + 1, 1, 1) - timedelta(days=1)
    else:
        end_date = datetime(year, month + 1, 1) - timedelta(days=1)
    
    # Belirlenen ay içindeki tüm günleri oluştur
    date_list = [start_date + timedelta(days=i) for i in range((end_date - start_date).days + 1)]
    
    new_data = []
    
    for date in date_list:
        # Günün hafta içi, cumartesi veya pazar olup olmadığını belirle
        if date.weekday() < 5:
            gun_tipi = "Hafta İçi"
        elif date.weekday() == 5:
            gun_tipi = "Cumartesi"
        else:
            gun_tipi = "Pazar"
        
        # Her gün için "G" ve "D" yönleri ayrı ayrı işlenir
        for yon in ["G", "D"]:
            # İlgili hat, ay ve gün tipine göre saat verilerini filtrele
            df_filtered = df_saatler[
                (df_saatler["Hat No"] == hat_no) & 
                (df_saatler["Ay"] == month) & 
                (df_saatler["Gün Tipi"] == gun_tipi) &
                (df_saatler["Yön"] == yon)
            ]
            
            if not df_filtered.empty:
                # Saat verilerini alırken ilk 4 kolondan sonraki sütunlar kullanılıyor
                saatler_raw = df_filtered.iloc[:, 4:].values.flatten()
                # Boş olmayan saatleri filtrele ve time objesi olarak al
                saatler = [s.time() if isinstance(s, datetime) else s for s in saatler_raw if pd.notna(s)]
            else:
                saatler = []
            
            # Mevcut satırdaki maksimum saat sütun sayısını belirle
            max_saat_sutun = max(len(df_saatler.columns[4:]), len(saatler))
            new_data.append([hat_no, date.strftime('%Y-%m-%d'), yon] + saatler + [""] * (max_saat_sutun - len(saatler)))
    
    # Yeni tabloyu oluştur: Kolon isimleri dinamik olarak oluşturuluyor
    column_names = ["Hat No", "Tarih", "Yön"] + [f"Saat{i+1}" for i in range(len(new_data[0]) - 3)]
    df_new = pd.DataFrame(new_data, columns=column_names)
    
    # Sonucu Excel dosyasına kaydet
    df_new.to_excel(output_path, index=False)
    print(f"Tablo başarıyla oluşturuldu: {output_path}")

# Kullanım örneği: Yıl bilgisi dışarıdan girilir, hat no ve ay dosyadan otomatik alınır.
generate_schedule("uploads/A1.xlsx", "Oluşturulan_Tablo.xlsx", year=2025)
