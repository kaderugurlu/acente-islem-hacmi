import os
import pandas as pd
from datetime import datetime, timedelta
import locale

# Yerel ayarları Türkçe olarak ayarla
locale.setlocale(locale.LC_TIME, 'tr_TR')

# Mevcut tarihi al
current_date = datetime.now()

# Bir önceki ayı hesapla
previous_month_date = current_date.replace(day=1) - timedelta(days=1)
previous_month = previous_month_date.strftime('%B').upper()  # Ay adını büyük harflerle al
previous_year = previous_month_date.year

# Yolun dinamik kısmını oluştur
path = f"Q:/_HiSenetl/İŞLEM DEFTERLERİ VE TESCİL/İŞLEM DEFTERLERİ/{previous_year}/{previous_month}"

# Dosyaları filtrele ve birleştir
csv_files = os.listdir(path)
oms_krep_combined_sum = pd.DataFrame()
diger_combined_sum = pd.DataFrame()

for file in csv_files:
    if "UID_" in file:
        df = pd.read_csv(f"{path}/{file}", delimiter=";", skiprows=[1])
        
        # İŞ BANKASI / YENİ OMS-KREP İŞLEMLERİ için filtreleme ve toplama
        filtered1 = df[df["REFERANS NO"].str.contains("7-", na=False)]
        oms_krep_sum = filtered1.groupby("ALIS_SATIS")["ISLEM HACMI"].sum().reset_index()
        oms_krep_combined_sum = pd.concat([oms_krep_combined_sum, oms_krep_sum], ignore_index=True)
        
        # İŞ BANKASI / DİĞER işlemleri için filtreleme ve toplama
        filtered2 = df[
            (df["REFERANS NO"].str.contains("1-", na=False)) &
            (df["HESAP TIPI"] == "M") &
            (df["HESAP NO"].apply(lambda x: len(str(x))) == 11)]
        diger_sum = filtered2.groupby("ALIS_SATIS")["ISLEM HACMI"].sum().reset_index()
        diger_combined_sum = pd.concat([diger_combined_sum, diger_sum], ignore_index=True)

# Tüm dosyalardan elde edilen toplamları tekrar grup olarak toplama
final_oms_krep_sum = oms_krep_combined_sum.groupby("ALIS_SATIS")["ISLEM HACMI"].sum().reset_index()
final_diger_sum = diger_combined_sum.groupby("ALIS_SATIS")["ISLEM HACMI"].sum().reset_index()

# Boş değerleri sıfır ile doldurma
final_oms_krep_sum = final_oms_krep_sum.fillna(0)
final_diger_sum = final_diger_sum.fillna(0)

# ALIS_SATIS sütunlarında eksik olan değerleri sıfır ile doldurma
all_alis_satis = pd.DataFrame({"ALIS_SATIS": ["A", "S"]})

final_oms_krep_sum = all_alis_satis.merge(final_oms_krep_sum, on="ALIS_SATIS", how="left").fillna(0)
final_diger_sum = all_alis_satis.merge(final_diger_sum, on="ALIS_SATIS", how="left").fillna(0)

# Birleştirilmiş DataFrame'leri yan yana gösterme
merged_df = pd.merge(final_oms_krep_sum, final_diger_sum, on="ALIS_SATIS", suffixes=("_OMS_KREP", "_DIGER"))

# Toplam satırını eklemek için toplamları hesaplayın
toplam_row = pd.DataFrame({
    "ALIS_SATIS": ["toplam"],
    "ISLEM HACMI_OMS_KREP": [merged_df["ISLEM HACMI_OMS_KREP"].sum()],
    "ISLEM HACMI_DIGER": [merged_df["ISLEM HACMI_DIGER"].sum()]})

# DataFrame'e toplam satırını ekleyin
merged_df = pd.concat([merged_df, toplam_row], ignore_index=True)

#Excel dosyasına yazdırma
output_file_name = f"Acente İşlem Hacmi_{previous_month}.xlsx"
merged_df.to_excel(output_file_name,index=False)