# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from geopy.geocoders import Nominatim
import time
import geopy.distance
import statsmodels.api as sm
from scipy.stats import chi2_contingency
import pyperclip
import win32com.client as win32
#adatok importálása
filepath = "data.xlsx"
data = pd.read_excel(filepath)

data.columns

#felesleges oszlopok eldobása
data_temp = data.copy()
data_temp.columns

data_temp = data_temp.drop(labels= ['Start time', 'Completion time', 'Email', 'Name', 'Last modified time'], axis=1)


#Ellentmondások keresése


data_temp.insert(loc=len(data_temp.columns),column="ellentmondas", value=np.where((data_temp["Jelenleg/alapvetően Budapesten laksz/tanulsz/dolgozol? (Legalább az egyik)"] == "Nem") & (data_temp["Hol laksz jelenleg/mi a tartózkodási helyed? (Megye/Főváros/Ország)"] == "Budapest"), True, False))


data_temp.query('ellentmondas == True')

#Nincs ellentmondás

#Nem Budapestiek eldobása
data_temp = data_temp.drop(data_temp[data_temp['Jelenleg/alapvetően Budapesten laksz/tanulsz/dolgozol? (Legalább az egyik)'] == "Nem"].index).reset_index(drop=True)



#hibás adatfelvétel javítása

data_temp["Melyik (külföldi) országban laksz jelenleg/mi a tartózkodási helyed?"] = np.where(data_temp['Hol laksz jelenleg/mi a tartózkodási helyed? (Megye/Főváros/Ország)'].isna(), "Budapest", data_temp["Melyik (külföldi) országban laksz jelenleg/mi a tartózkodási helyed?"])

data_temp["Hol laksz jelenleg/mi a tartózkodási helyed? (Megye/Főváros/Ország)"] = np.where(data_temp['Hol laksz jelenleg/mi a tartózkodási helyed? (Megye/Főváros/Ország)'].isna(), "Budapest", data_temp["Hol laksz jelenleg/mi a tartózkodási helyed? (Megye/Főváros/Ország)"])


#Városok egy oszlopba rendezése
data_temp
data_temp['szarmazas'] = data_temp.iloc[:, 4:25].stack().dropna().reset_index(drop=True)
data_temp = data_temp.drop(data_temp.iloc[:, 4:25],axis = 1)
data_temp['jelenlegi'] = data_temp.iloc[:, 5:26].stack().dropna().reset_index(drop=True)
data_temp = data_temp.drop(data_temp.iloc[:, 5:26],axis = 1)



# #VÁROSOK KOORDINÁTÁINAK MEGSZERZÉSE ÉS KIÍRÁSA EGY EXCEL FILEBA
# #CSAK AKKOR KELL ÚJRA LEFUTTATNI, HA ÚJ VÁROS VAN

# #Minden eddigi város és helyszín
# city_list = np.array(data_temp['szarmazas'])
# city_list = pd.Series(np.append(city_list, data_temp['jelenlegi']))
# city_list = city_list.unique()


# # Üres numpy array az új oszlopoknak
# coordinates = np.empty((len(city_list), 2), dtype=float)

# # Geopy geokódoló objektum létrehozása
# geolocator = Nominatim(user_agent="city_coordinates")

# # Késleltetés másodpercekben a lekérdezések között
# delay = 2

# # Városok koordinátáinak lekérdezése és hozzáadása az array-hez
# for i, city in enumerate(city_list):
#     location = geolocator.geocode(city)
#     if location:
#         coordinates[i] = (location.latitude, location.longitude)
#     else:
#         print(f"Could not find coordinates for {city}")
#     time.sleep(delay)  # Várakozás a következő lekérdezés előtt

# # Az eredeti városlista és az új koordináta-oszlopok hozzáadása
# city_list_with_coordinates = np.column_stack((city_list, coordinates))

# print(city_list_with_coordinates)

# pd.DataFrame(city_list_with_coordinates).to_excel("city_coordinates.xlsx")


coords = pd.read_excel("city_coordinates.xlsx")

#város koordináták hozzáadása

data_temp = data_temp.merge(coords, how="left", left_on="szarmazas", right_on=0)
data_temp = data_temp.drop(["szarmazas", "Unnamed: 0", 0], axis=1)
data_temp = data_temp.rename(columns={"lat":"szarmazas_lat","lon":"szarmazas_lon"})

data_temp = data_temp.merge(coords, how="left", left_on="jelenlegi", right_on=0)
data_temp = data_temp.drop(["jelenlegi", "Unnamed: 0", 0], axis=1)
data_temp = data_temp.rename(columns={"lat":"jelenlegi_lat","lon":"jelenlegi_lon"})

data_temp.groupby(by=["Milyen területen dolgozol/tanulsz? (Ha dolgozol és tanulsz is, de nem egy terület, aszerint válaszolj, amelyik jobban leír téged.)"]).count()


#kategóriák újraírása

categories = {
    "Agrár" : "Mezőgazdaság és ipar",
    "Ipar" : "Mezőgazdaság és ipar",
    "Feldolgozóipar" : "Mezőgazdaság és ipar",
    "Szállítás, raktározás, posta, távközlés" : "Mezőgazdaság és ipar",
    "Logisztika" : "Mezőgazdaság és ipar",
    "Bölcsészettudomány" : "Bölcsészettudomány",
    "Gazdaságtudományok" : "Gazdaságtudományok",
    "Informatika" : "Informatika",
    "Gazdaságinformatika" : "Informatika",
    "Jogi" : "Jogi",
    "Biztosítás" : "Jogi",
    "Államtudományi" : "Jogi",
    "Oktatás" : "Oktatás",
    "Pedagógusképzés" : "Oktatás",
    "Gimnázium" : "Középiskolai tanulmányok",
    "Gimnázium " : "Középiskolai tanulmányok",
    "Általános gimnáziumi műveltség" : "Középiskolai tanulmányok",
    "Középiskola" : "Középiskolai tanulmányok",
    "Középiskolai tanulmányok" : "Középiskolai tanulmányok",
    "Orvos- és egészségtudomány" : "Orvos- és egészségtudomány",
    "Egészségügyi, szociális ellátás" : "Orvos- és egészségtudomány",
    "Egyéb közösségi, személyi szolgáltatás" : "Szolgáltatások",
    "Kereskedelem, javítás" : "Szolgáltatások",
    "Pénzügyi közvetítés" : "Szolgáltatások",
    "HR - munkaerőközvetítő" : "Szolgáltatások",
    "Műszaki" : "Műszaki",
    "Művészet" : "Művészet",
    "Rendészettudomány" : "Rendészettudomány",
    "Rendvèdelem" : "Rendészettudomány",
    "Sporttudomány" : "Sporttudomány",
    "Turizmus" : "Turizmus",
    "Szálláshely-szolgáltatás, vendéglátás" : "Turizmus",
    "Természettudomány" : "Természettudomány",
    "Társadalomtudomány" : "Társadalomtudomány",
    "Területen kívüli szervezet" : "Egyéb",
    "Nem szeretném megadni" : "Egyéb"}

categories = pd.DataFrame(categories.items(), columns=["régi_terület","terület"])

data_temp = data_temp.merge(categories, how="left", left_on="Milyen területen dolgozol/tanulsz? (Ha dolgozol és tanulsz is, de nem egy terület, aszerint válaszolj, amelyik jobban leír téged.)", right_on="régi_terület")

data_temp = data_temp.drop(labels=["Milyen területen dolgozol/tanulsz? (Ha dolgozol és tanulsz is, de nem egy terület, aszerint válaszolj, amelyik jobban leír téged.)","Ha nem tudsz dönteni, akkor add meg röviden mit tanulsz/milyen munkakörben dolgozol.","régi_terület"], axis=1)


#otthontól való távolság kiszámítása

for index, row in data_temp.iterrows():
    origin = (row["szarmazas_lat"], row["szarmazas_lon"])
    destination = (row["jelenlegi_lat"], row["jelenlegi_lon"])
    distance = geopy.distance.geodesic(origin, destination).kilometers
    data_temp.at[index, "tavolsag_otthontol"] = distance
    

#folyamatban lévő/befejezett tanulmányok kitöltéseinek javítása

data_temp['Ha már nem tanulsz, mi a legmagasabb végzettséged?'] = np.where(data_temp['Ha jelenleg tanulsz, milyen szintű képzést végzel?'] != "Nem tanulok", "Még tanulok", data_temp['Ha már nem tanulsz, mi a legmagasabb végzettséged?'])

#felesleges oszlopok eldobása

data_temp = data_temp.drop(labels=["ID","Jelenleg/alapvetően Budapesten laksz/tanulsz/dolgozol? (Legalább az egyik)", "ellentmondas"], axis=1)

#idősek eldobása
data_temp = data_temp.drop(data_temp[data_temp["Hány éves vagy?"] > 25].index).reset_index(drop=True)

#oszlopok átnevezése

data_temp = data_temp.rename(columns={"Hány éves vagy?":"kor",data_temp.columns[1]:"hol_nott_fel",data_temp.columns[2]:"jelenlegi_lakhely","Jelenleg is otthon, a családi lakhelyen laksz? (Tehát nincsen más lakhely, amit tartózkodási helynek neveznél.)":"otthon_lakik","Milyen típusú ingatlan a tartózkodási helyed?":"otthon_tipus","Amennyiben külön laksz, magadnak fizeted a lakhelyed költségeit?":"fizet_lakhatasert","Milyen nemű vagy?":"nem","Ha jelenleg tanulsz, milyen szintű képzést végzel?":"jelenlegi_tanulmany","Ha már nem tanulsz, mi a legmagasabb végzettséged?":"legmagasabb_vegzettseg","Jelenleg dolgozol?":"dolgozik","Melyik egyetemen tanulsz? (Ha dolgozol/nem egyetemen tanulsz/nem találod a listában az egyetemedet, akkor értelemszerűen a megfelelő opciót válaszd.)":"egyetem","Havonta átlagosan nettó mennyi pénzt keresel? (Munka, ösztöndíj stb.)":"sajat_kereset","Havonta átlagosan mennyi pénzt kapsz szüleidtől?":"kapott_penz","A szüleid pénzen kívül rendszeresen kisegítenek mással?":"egyeb_segitseg","Van bankszámlád?":"bankszamla","Van bankkártyád?":"bankkartya","Hallottál már az okoseszközzel való fizetésről?":"tud_okoseszkozrol","Van megfelelő eszközöd és kompatibilis bankkártyád ahhoz, hogy legyen lehetőséged okoseszközzel fizetni?":"van_okoseszkoz","Hogyan viszonyulsz a technológiához?":"technologia_viszony","Hallottál arról, hogy 2021. január 1-től a kereskedelemről szóló 2005. évi CLXIV. törvény módosítása értelmében minden online pénztárgép mellett kötelező biztosítani a bankkártyás fizetés lehetősé...":"hallott_torvenyrol","Milyen gyakran szoktál fizikai boltokban vásárolni/nem online fizetni?":"milyen_gyakran_fizet","Milyen gyakran jársz szórakozóhelyekre?":"milyen_gyakran_szorakozik","Milyen gyakran jársz nyaralni külföldre saját pénzből?":"milyen_gyakran_nyaral","Szoktál gyakran olyan helyekre járni, ahol csak bankkártyával lehet fizetni?":"csak_bankkartyas_hely"," Kerültél már olyan helyzetbe az elmúlt 3 évben, hogy nem tudtál készpénzzel fizetni?":"nem_kp_fizetes"," Kerültél már olyan helyzetbe az elmúlt 3 évben, hogy nem tudtál bankkártyával fizetni?":"nem_bankkartya_fizetes","Kértek már meg arra, hogy amennyiben módodban áll, inkább készpénzzel fizess?":"inkabb_kp","Szoktál különösen aggódni amiatt, hogy ellopják a nálad tartott készpénzt?":"kp_ellop","Szoktál különösen aggódni amiatt, hogy ellopják a bankkártyádat?":"kartya_ellop","Szoktál különösen aggódni amiatt, hogy lemerül a fizetéshez használt okoseszközöd, és így nem tudsz fizetni? ":"okoseszkoz_lemerul","Kérlek állítsd sorrendbe, hogy amikor fizikai helyen fizetsz, milyen módon szeretsz fizetni, mit preferálsz!":"preferencia"})

#kor stringként kezelése, hogy kategorikus változóként kezelje a dummy függvény

data_temp["kor"] = data_temp["kor"].astype(str)

#Szülői segítség manuális dummizása

szuloi_segitseg = pd.DataFrame()

szuloi_segitseg["elofizetes"] = np.where(data_temp["egyeb_segitseg"].str.contains("előfizetést fizetnek"), True, False)

szuloi_segitseg["etel"] = np.where(data_temp["egyeb_segitseg"].str.contains("Étellel"), True, False)

szuloi_segitseg["lakhatas"] = np.where(data_temp["egyeb_segitseg"].str.contains("A családi lakhelyen kívüli lakhatást biztosítják valamely módon"), True, False)


#Megyék kategóriába rendezése

Alfold_es_eszak = ("Borsod-Abaúj-Zemplén Vármegye","Heves Vármegye","Nógrád Vármegye", "Hajdú-Bihar Vármegye", "Jász-Nagykun-Szolnok Vármegye", "Szabolcs-Szatmár-Bereg Vármegye", "Bács-Kiskun Vármegye", "Békés Vármegye", "Csongrád-Csanád Vármegye")
Kozep_magyarorszag = ("Pest Vármegye", "Budapest")
Dunantul = ("Fejér Vármegye", "Komárom-Esztergom Vármegye", "Veszprém Vármegye", "Baranya Vármegye", "Somogy Vármegye", "Tolna Vármegye", "Győr-Moson-Sopron Vármegye", "Vas Vármegye", "Zala Vármegye")

data_temp["hol_nott_fel"] = np.where(data_temp["hol_nott_fel"].isin(Alfold_es_eszak),"Alföld és Észak", data_temp["hol_nott_fel"])

data_temp["hol_nott_fel"] = np.where(data_temp["hol_nott_fel"].isin(Kozep_magyarorszag),"Középmagyarország", data_temp["hol_nott_fel"])

data_temp["hol_nott_fel"] = np.where(data_temp["hol_nott_fel"].isin(Dunantul),"Dunántúl", data_temp["hol_nott_fel"])

data_temp["jelenlegi_lakhely"] = np.where(data_temp["jelenlegi_lakhely"].isin(Alfold_es_eszak),"Alföld és Észak", data_temp["jelenlegi_lakhely"])

data_temp["jelenlegi_lakhely"] = np.where(data_temp["jelenlegi_lakhely"].isin(Kozep_magyarorszag),"Középmagyarország", data_temp["jelenlegi_lakhely"])

data_temp["jelenlegi_lakhely"] = np.where(data_temp["jelenlegi_lakhely"].isin(Dunantul),"Dunántúl", data_temp["jelenlegi_lakhely"])

#egy darab egyedi kitöltés eldobása
data_temp = data_temp.drop(data_temp[data_temp['nem'] == "NembináristranszpoliamoruszluxusTnyelürotációskapa"].index).reset_index(drop=True)

#preferencia szerint eldobás
data_temp.groupby('preferencia').agg(count=('preferencia', 'count'))

data_temp = data_temp.drop(data_temp[data_temp['preferencia'] == "Készpénz;Okoszeszköz;Bankkártya;"].index).reset_index(drop=True)

#nem fizet bizonyos módon átírása "nem"-re
data_temp["kp_ellop"].replace("Nem hordok magamnál készpénzt","Nem", inplace=True)
data_temp["kartya_ellop"].replace("Nem hordok magamnál bankkártyát / Nincs bankkártyám", "Nem", inplace=True)
data_temp["okoseszkoz_lemerul"].replace("Nem szoktam így fizetni", "Nem", inplace=True)

#tisztított adatok exportálása
data_temp.to_excel("clean_data.xlsx")

#adatok dummyzása
X = pd.concat([pd.get_dummies(data_temp.iloc[:, list(range(13)) + list(range(14, 30)) + [35]], drop_first=False).astype(int),szuloi_segitseg.astype(int) ,data_temp.iloc[:,[31,32,33,34,36]]], axis=1,).drop({202,203})


#dummy oszlopokból referencia eldobása, plusz egyéb érdekességként feltett kérdések eldobása
Columns = X.columns
X.drop(columns={"kor_18","hol_nott_fel_Külföld","jelenlegi_lakhely_Külföld","otthon_lakik_Igen","otthon_tipus_Családi közös lakóhely",
"fizet_lakhatasert_Nem",
"nem_Nő",
"jelenlegi_tanulmany_Nem tanulok",
"legmagasabb_vegzettseg_Még tanulok",
"dolgozik_Nem",
"egyetem_Nem egyetemen tanulok",
"sajat_kereset_0Ft",
"kapott_penz_0Ft",
"bankszamla_Nem",
"bankszamla_Igen",
"bankkartya_Nem",
"bankkartya_Igen",
"tud_okoseszkozrol_Igen",
"van_okoseszkoz_Nem ",
"van_okoseszkoz_Igen",
"van_okoseszkoz_Nem tudom",
"technologia_viszony_Győlölöm",
"hallott_torvenyrol_Nem",
"milyen_gyakran_fizet_Ritkán/Nem szoktam",
"milyen_gyakran_szorakozik_Nem szoktam / maximum évente egyszer",
"milyen_gyakran_nyaral_Nem szoktam / A szüleimmel megyek, és ők szoktak fizetni",
"csak_bankkartyas_hely_Nem",
"nem_kp_fizetes_Nem",
"nem_bankkartya_fizetes_Nem",
"inkabb_kp_Nem",
"kp_ellop_Nem",
"kartya_ellop_Nem",
"okoseszkoz_lemerul_Nem",
"terület_Egyéb",
}, inplace=True)

#1 vagy -1 korrelációjú oszlopok esetén az egyik eldobása

correlation_matrix = X.corr()
#correlation_matrix.to_excel("Correlation_matrix.xlsx")
X.drop(columns={"jelenlegi_tanulmany_Gimnázium", "egyetem_Magyar Testnevelési és Sporttudományi Egyetem"}, inplace=True, axis=1)

#kevés különböző értékkel rendelkező oszlop eldobása

column_sums = X.sum()
columns_to_drop = column_sums[column_sums < 11].index
X = X.drop(columns=columns_to_drop)

#Célváltozó létrehozása

y = data_temp["preferencia"]
y = y.replace(";","", regex=True)

#Kategória átnevezése, hogy az legyen a referencia kategória
y = y.replace("OkoszeszközBankkártyaKészpénz","AOkoszeszközBankkártyaKészpénz", regex=True)



#magyarázó változók szűrése


chi2_results = []

# Chi-négyzet próba elvégzése minden bináris változóra
for binary_var in X:
    # Keresztezési tábla létrehozása
    contingency_table = pd.crosstab(y, X[binary_var])
    # Chi-négyzet próba végrehajtása
    chi2, p, dof, expected = chi2_contingency(contingency_table)
    # Eredmény hozzáadása a listához
    chi2_results.append({'Binary Variable': binary_var,
                          'Chi-square': chi2,
                          'P-value': p,
                          'Degrees of Freedom': dof,
                          'Expected Frequencies': expected})

chi2_results = pd.DataFrame(chi2_results)

#20 legkisebb p-értékű változó kiválasztása
n20 = chi2_results.nsmallest(20, "P-value")

X = X.loc[:,n20.iloc[:,0]]

#Oszlopnevek tisztítása
X.columns = X.columns.str.replace(' ', '_')
X.columns = X.columns.str.replace(',', '')


#Konstans tag hozzáadása
X = sm.add_constant(X)

#A későbbi AIC szűrés alapján változók eldobása
X.drop(columns={"hol_nott_fel_Alföld_és_Észak","otthon_tipus_Kollégium","szarmazas_lon","csak_bankkartyas_hely_Igen","terület_Bölcsészettudomány","szarmazas_lat","technologia_viszony_Közömbös ","milyen_gyakran_nyaral_Évente_egyszer"}, inplace=True )

#modell felállítsa statsmodels segítségével
model = sm.MNLogit(y, X)
result = model.fit_regularized(method='l1')

#modell eredményeinek kiírása

print(result.summary())
print("BIC érték:", result.bic)
print("AIC érték:", result.aic)


# A summary szövegének elmentése
summary_str = result.summary().as_text()

# Szöveg másolása a vágólapra
pyperclip.copy(summary_str)

# Excel alkalmazás megnyitása és aktív munkafüzet beillesztése
excel = win32.Dispatch("Excel.Application")
excel.Visible = True  # Megnyitja az Excel alkalmazást

# Új munkafüzet hozzáadása
workbook = excel.Workbooks.Add()
sheet = workbook.Sheets(1)

# Az első cellába való beillesztés
sheet.Range("A1").Select()
sheet.Paste()

file_name = "D:\OneDrive - Corvinus University of Budapest\Egyetem hivatalos\TDK\python\summary_output.xlsx"
workbook.SaveAs(file_name)

# AIC és BIC változószűrés

X_j = X.copy()
szures_sorrendje = []
szamlalo=1

while X_j.columns.size > 2:
    szelekcio = []
    
    for i in range(X_j.columns.size):
        X_i = X_j.copy()
        jelenlegi_valtozo = X_i.columns[i]
        print(jelenlegi_valtozo)
        X_i.drop(columns={jelenlegi_valtozo}, inplace=True )
    
        model = sm.MNLogit(y, X_i)
        #result = model.fit()
        result = model.fit_regularized(method='l1')
        #result = model.fit(maxiter=100)
    
        szelekcio.append({"valtozo":jelenlegi_valtozo,
                          "BIC": result.bic,
                          "AIC":result.aic,
                          "R2":result.prsquared,
                          "logl":result.llf})
        print(i)
    szelekcio = pd.DataFrame(szelekcio)
    legrosszabb_valtozo = szelekcio.nsmallest(2, "AIC").reset_index()
    if (legrosszabb_valtozo.at[0, "valtozo"] == "const"):
        legrosszabb_valtozo.drop(index=0, inplace=True)
    szures_sorrendje.append({"sorszam":szamlalo,
                              "valtozo":legrosszabb_valtozo.at[0,"valtozo"],
                              "AIC":legrosszabb_valtozo.at[0, "AIC"],
                              "R2":legrosszabb_valtozo.at[0, "R2"],
                              "logl":legrosszabb_valtozo.at[0, "logl"]})
    X_j.drop(columns={legrosszabb_valtozo.at[0, "valtozo"]}, inplace=True)
    print(X_j.columns.size)
    szamlalo += 1


szures_sorrendje_BIC = pd.DataFrame(szures_sorrendje)

X_j = X.copy()
szures_sorrendje = []
szamlalo=1

while X_j.columns.size > 2:
    szelekcio = []
    
    for i in range(X_j.columns.size):
        X_i = X_j.copy()
        jelenlegi_valtozo = X_i.columns[i]
        print(jelenlegi_valtozo)
        X_i.drop(columns={jelenlegi_valtozo}, inplace=True )
    
        model = sm.MNLogit(y, X_i)
        #result = model.fit()
        result = model.fit_regularized(method='l1')
        #result = model.fit(maxiter=100)
    
        szelekcio.append({"valtozo":jelenlegi_valtozo,
                          "BIC": result.bic,
                          "AIC":result.aic,
                          "R2":result.prsquared,
                          "logl":result.llf})
        print(i)
    szelekcio = pd.DataFrame(szelekcio)
    legrosszabb_valtozo = szelekcio.nsmallest(2, "AIC").reset_index()
    if (legrosszabb_valtozo.at[0, "valtozo"] == "const"):
        legrosszabb_valtozo.drop(index=0, inplace=True)
    szures_sorrendje.append({"sorszam":szamlalo,
                              "valtozo":legrosszabb_valtozo.at[0,"valtozo"],
                              "AIC":legrosszabb_valtozo.at[0, "AIC"],
                              "R2":legrosszabb_valtozo.at[0, "R2"],
                              "logl":legrosszabb_valtozo.at[0, "logl"]})
    X_j.drop(columns={legrosszabb_valtozo.at[0, "valtozo"]}, inplace=True)
    print(X_j.columns.size)
    szamlalo += 1

szures_sorrendje_AIC = pd.DataFrame(szures_sorrendje)


with pd.ExcelWriter("szures_sorrendje.xlsx") as writer:
    szures_sorrendje_BIC.to_excel(writer, sheet_name="BIC", index=False)
    szures_sorrendje_AIC.to_excel(writer, sheet_name="AIC", index=False)
    
