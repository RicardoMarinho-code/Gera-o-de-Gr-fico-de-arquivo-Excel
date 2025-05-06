import os
import pandas as pd
import matplotlib.pyplot as plt

script_dir = os.path.dirname(os.path.abspath(__file__))

#vê se o arquivo está na mesma pasta que o script
file_path = os.path.join(script_dir, "base_ogu_202503_snh.xlsx")

base_ogu = pd.read_excel(file_path, sheet_name="base_ogu")
base_ogu["Unidades Contratadas"] = pd.to_numeric(base_ogu["Unidades Contratadas"], errors="coerce")

base_ogu["Ano"] = pd.to_numeric(base_ogu["Ano"], errors="coerce").astype('Int64')
eita = base_ogu.groupby("Ano")["Unidades Contratadas"].sum().reset_index()
eita.columns = ["Rótulos de Linha", "Soma de mcmv_ogu_22_qtd_uh"]
eita = eita.dropna(subset=["Soma de mcmv_ogu_22_qtd_uh", "Rótulos de Linha"])

eita = eita.sort_values(by="Rótulos de Linha", ascending=True)
eita.plot(x="Rótulos de Linha", y="Soma de mcmv_ogu_22_qtd_uh", kind="bar", figsize=(10, 6))
plt.title("Tabela de Unidades Contratadas por Ano")
plt.xlabel("Ano")
plt.ylabel("Unidades Contratadas")

plt.savefig(os.path.join(script_dir, "Tabela.png"))

# Salva o DataFrame como CSV
eita.to_csv(os.path.join(script_dir, "Tabela.csv"), index=True)

plt.show()