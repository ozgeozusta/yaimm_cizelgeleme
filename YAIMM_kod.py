from gurobipy import Model, GRB, quicksum
import pandas as pd

# Parametreler
personel_sayisi = 50
gun_sayisi = 5
vardiya_sayisi = 2
gorev_sayisi = 8

# İhtiyac duyulan personel sayisi (M_jk) [gun, vardiya, gorev]
M = [[[0 for _ in range(gorev_sayisi)] for _ in range(vardiya_sayisi)] for _ in range(gun_sayisi)]

# Tabloya göre M degerleri girilmektedir
# Gun 1
M[0][0] = [2, 3, 3, 2, 0, 1, 3, 0]  # Vardiya 1
M[0][1] = [2, 3, 3, 1, 0, 1, 2, 0]  # Vardiya 2
# Gun 2
M[1][0] = [2, 3, 3, 2, 0, 1, 3, 0]  # Vardiya 1
M[1][1] = [2, 3, 3, 1, 0, 1, 2, 0]  # Vardiya 2
# Gun 3
M[2][0] = [5, 6, 3, 2, 6, 5, 5, 0]  # Vardiya 1
M[2][1] = [4, 6, 3, 1, 6, 5, 5, 3]  # Vardiya 2
# Gun 4
M[3][0] = [5, 6, 3, 3, 6, 5, 5, 0]  # Vardiya 1
M[3][1] = [5, 6, 3, 3, 6, 5, 5, 3]  # Vardiya 2
# Gun 5
M[4][0] = [5, 6, 3, 3, 6, 5, 5, 0]  # Vardiya 1
M[4][1] = [5, 6, 3, 3, 6, 5, 5, 6]  # Vardiya 2

# Yetenek durumu (Y_il): Tüm personeller tum görevlere atanabilir
Y = [[1 for l in range(gorev_sayisi)] for i in range(personel_sayisi)]

# Model
model = Model("Personel Atama")

# Karar Degiskenleri
x = model.addVars(personel_sayisi, gun_sayisi, vardiya_sayisi, gorev_sayisi, vtype=GRB.BINARY, name="x")

# Amaç Fonksiyonu
model.setObjective(quicksum(x[i,j,k,l] * Y[i][l] for i in range(personel_sayisi)
                                                   for j in range(gun_sayisi)
                                                   for k in range(vardiya_sayisi)
                                                   for l in range(gorev_sayisi)), GRB.MAXIMIZE)


# Kısıtlar
# 1. Her görev için ihtiyaç duyulan personel atanmalı
for j in range(gun_sayisi):
    for k in range(vardiya_sayisi):
        for l in range(gorev_sayisi):
            model.addConstr(quicksum(x[i,j,k,l] for i in range(personel_sayisi)) == M[j][k][l], f"gorev_atama_{j}_{k}_{l}")

# 2. Personel yetkinliklerine göre atanacak
for i in range(personel_sayisi):
    for l in range(gorev_sayisi):
        model.addConstr(quicksum(x[i,j,k,l] for j in range(gun_sayisi) 
                                             for k in range(vardiya_sayisi)) <= Y[i][l], f"yetenek_{i}_{l}")

# 3. Bir personel bir vardiyada yalnızca bir göreve atanabilir
for i in range(personel_sayisi):
    for j in range(gun_sayisi):
        for k in range(vardiya_sayisi):
            model.addConstr(quicksum(x[i,j,k,l] for l in range(gorev_sayisi)) <= 1, f"tek_gorev_{i}_{j}_{k}")

# 4. Protokol ve VIP Hizmetleri (Görev 5), 3,4,5 gün tüm vardiyalarda
for j in [2, 3, 4]:
    for k in range(vardiya_sayisi):
        model.addConstr(quicksum(x[i,j,k,4] for i in range(personel_sayisi)) >= M[j][k][4], f"protokol_{j}_{k}")

# 5. Geri Bildirim ve Raporlama (Görev 8), 3,4,5 gün sadece 2. vardiya
for j in [2, 3, 4]:
    model.addConstr(quicksum(x[i,j,1,7] for i in range(personel_sayisi)) >= M[j][1][7], f"geri_bildirim_{j}")

# Modelin çözülmesi
model.optimize()

# Çözüm Sonuçları ve Excel Çıktısı
if model.status == GRB.OPTIMAL:
    
    # Excel verisi için liste
    excel_data = []
    for i in range(personel_sayisi):
        row = [i + 1]  # Personel numarası
        for j in range(gun_sayisi):
            for k in range(vardiya_sayisi):
                assigned = ""  # Atanma durumu
                for l in range(gorev_sayisi):
                    if x[i, j, k, l].x > 0.5:
                        assigned = f"G{l+1}"  # Görev numarası
                row.append(assigned)
        excel_data.append(row)

    # DataFrame'e dönüştür
    columns = ["Personel"] + [f"{j+1}. Gün - Vardiya {k+1}" for j in range(gun_sayisi) for k in range(vardiya_sayisi)]
    df = pd.DataFrame(excel_data, columns=columns)

    # Excel dosyasına yaz
    df.to_excel("personel_atama.xlsx", index=False)
    print("Excel dosyası 'personel_atama.xlsx' olarak kaydedildi.")
else:
    print("Optimal çözüm bulunamadi.")
