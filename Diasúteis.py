from datetime import date
from workalendar.america import Brazil
import pandas as pd

# Estados por cidade/porto
cidades_estados = {
    "Manaus": "AM",
    "Vila do Conde": "PA",
    "Pecem": "CE",
    "Suape": "PE",
    "Salvador": "BA",
    "Vitoria": "ES",
    "Rio": "RJ",
    "Santos": "SP",
    "Itapoa": "SC",
    "Imbituba": "SC",
    "Rio Grande": "RS",
    "Itajai": "SC",
    "Paranagua": "PR"
}

# Datas
inicio = date(2025, 6, 1)
fim = date(2025, 6, 30)

# Resultado
resultado = []
for cidade, uf in cidades_estados.items():
    cal = Brazil(state=uf)
    dias_uteis = cal.get_working_days_delta(inicio, fim)  # Corrigido aqui!
    resultado.append([cidade, dias_uteis])

# Exibir
df = pd.DataFrame(resultado, columns=["Cidade/Porto", "Dias Úteis em Junho 2025"])
print(df.sort_values("Cidade/Porto"))

