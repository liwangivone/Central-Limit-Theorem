import pandas as pd
import numpy as np
import random

# Baca file Excel
file_path = "CLT_Datasets.xlsx"
df = pd.read_excel(file_path, sheet_name=None)  # baca semua sheet
data_dict = {name: sheet['Data'].dropna().values for name, sheet in df.items()}  # ambil nilai kolom 'Data'

# Fungsi untuk sampling
def clt_analysis(data, sample_sizes=[1, 5, 10, 20], num_groups=1000):
    results = []
    population_mean = np.mean(data)

    for size in sample_sizes:
        group_means = []
        for _ in range(num_groups):
            sample = random.choices(data, k=size)
            sample_mean = np.mean(sample)
            group_means.append(sample_mean)

        mean_of_means = np.mean(group_means)
        selisih = abs(mean_of_means - population_mean)

        # Rataan tiap sampel hanya ditampilkan saat 1 kelompok
        rata_rata_tiap_sampel = group_means[0] if num_groups == 1 else "Beragam"

        results.append({
            "Jumlah Kelompok Sampel": size,
            "Rataan Tiap Sampel": rata_rata_tiap_sampel,
            "Rataan dari Rataan Sampel": round(mean_of_means, 2),
            "Rataan Populasi": round(population_mean, 2),
            "Selisih": round(selisih, 2)
        })

    return results

# Proses untuk semua dataset
final_results = []

for data_name, data_values in data_dict.items():
    print(f"Processing {data_name}...")
    result = clt_analysis(data_values, sample_sizes=[1, 5, 10, 20], num_groups=1000)
    for row in result:
        row["Dataset"] = data_name
        final_results.append(row)

# Buat DataFrame hasil
result_df = pd.DataFrame(final_results)
# Susun urutan kolom
result_df = result_df[["Dataset", "Jumlah Kelompok Sampel", "Rataan Tiap Sampel", 
                       "Rataan dari Rataan Sampel", "Rataan Populasi", "Selisih"]]

# Tampilkan
print(result_df)

# Simpan ke Excel jika diinginkan
# result_df.to_excel("hasil_clt.xlsx", index=False)
