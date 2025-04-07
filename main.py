
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Ler Excel e Calcular Tempo Médio</title>
    <script src="https://cdn.jsdelivr.net/pyodide/v0.23.4/full/pyodide.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
            background-color: #f4f4f4;
        }
        .header {
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 20px;
        }
        .logo {
            max-width: 150px; /* Adjust size as needed */
        }
        .controls {
            display: flex;
            align-items: center;
            gap: 10px;
        }
        button {
            padding: 5px 5px;
            background-color: #040493;
            color: white;
            border: none;
            cursor: pointer;
        }
        button:hover {
            background-color: #45a049;
        }
        #output {
            width: 100%;
            height: 400px;
            border: 1px solid #ccc;
            padding: 10px;
            background-color: white;
            overflow-y: scroll;
            white-space: pre-wrap;
        }
    </style>
</head>
<body> 
    <div class="header">
        <img src="/img/LOGO_AZUL_LINHAS_AEREAS.png" alt="Logo" class="logo">
        <div class="controls">
            <input type="file" id="excelFile" accept=".xlsx, .xls" />
            <button onclick="processarExcel()">Processar Arquivo</button>
        </div>
    </div>
    <div id="output">Selecione um arquivo Excel e clique em "Processar Arquivo"...</div>

    <script>
        async function loadPyodideAndRun() {
            let pyodide = await loadPyodide();
            await pyodide.loadPackage("pandas");
            return pyodide;
        }

        let pyodideReady = loadPyodideAndRun();

        async function processarExcel() {
            try {
                let pyodide = await pyodideReady;

                const fileInput = document.getElementById("excelFile");
                if (!fileInput.files.length) {
                    document.getElementById("output").innerText = "Por favor, selecione um arquivo Excel.";
                    return;
                }

                const file = fileInput.files[0];
                const reader = new FileReader();

                reader.onload = async function (e) {
                    try {
                        const data = new Uint8Array(e.target.result);
                        const workbook = XLSX.read(data, { type: "array", cellDates: true });
                        const firstSheet = workbook.SheetNames[0];
                        const worksheet = workbook.Sheets[firstSheet];

                        const jsonData = XLSX.utils.sheet_to_json(worksheet, { raw: false, dateNF: "yyyy-mm-dd hh:mm:ss" });

                        if (!jsonData[0] || !jsonData[0].AWB || !jsonData[0].Chegada || !jsonData[0].Retirada) {
                            document.getElementById("output").innerText = 
                                "O arquivo Excel deve conter as colunas 'AWB', 'Chegada' e 'Retirada'.";
                            return;
                        }

                        const dadosJson = JSON.stringify(jsonData);
                        pyodide.globals.set("dados_json", dadosJson);

                        let pythonCode = `
import pandas as pd
import json

dados = json.loads(dados_json)
df = pd.DataFrame(dados)
df = df.drop_duplicates(subset=['AWB'], keep='first')

def parse_date(date_str):
    if pd.isna(date_str) or date_str is None:
        return pd.NaT
    date_str = str(date_str).strip()
    if not date_str:
        return pd.NaT
    try:
        return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M:%S')
    except ValueError:
        try:
            return pd.to_datetime(date_str, format='%d/%m/%Y %H:%M')
        except ValueError:
            try:
                return pd.to_datetime(date_str, format='%Y-%m-%d %H:%M')
            except ValueError:
                try:
                    return pd.to_datetime(date_str)
                except ValueError:
                    return pd.NaT

df['Chegada'] = df['Chegada'].apply(parse_date)
df['Retirada'] = df['Retirada'].apply(parse_date)
df = df.dropna(subset=['Chegada', 'Retirada'])

if df.empty:
    raise ValueError("Nenhuma linha válida após a conversão de datas. Verifique o formato das datas no arquivo Excel.")

df['Tempo_H'] = (df['Retirada'] - df['Chegada']).dt.total_seconds() / 3600
df = df[df['Tempo_H'] >= 0]
tempo_medio_geral = df['Tempo_H'].mean()

output = "Detalhes das Cargas:\\n"
output += f"{'AWB':<10} {'Chegada':<25} {'Retirada':<25} {'Tempo (h)':<10}\\n"
output += "-" * 70 + "\\n"
for _, row in df.iterrows():
    chegada_str = row['Chegada'].strftime('%Y-%m-%d %H:%M:%S') if pd.notna(row['Chegada']) else 'N/A'
    retirada_str = row['Retirada'].strftime('%Y-%m-%d %H:%M:%S') if pd.notna(row['Retirada']) else 'N/A'
    output += f"{str(row['AWB']):<10} {chegada_str:<25} {retirada_str:<25} {row['Tempo_H']:<10.2f}\\n"

output += f"\\nTempo Médio de Processamento das Cargas: {tempo_medio_geral:.2f} horas\\n"
output
`;

                        let result = await pyodide.runPythonAsync(pythonCode);
                        document.getElementById("output").innerText = result;
                    } catch (error) {
                        console.error("Erro ao processar o arquivo Excel:", error);
                        document.getElementById("output").innerText = "Erro ao processar o arquivo Excel: " + error.message;
                    }
                };

                reader.onerror = function () {
                    document.getElementById("output").innerText = "Erro ao ler o arquivo Excel.";
                };

                reader.readAsArrayBuffer(file);
            } catch (error) {
                console.error("Erro ao processar os dados:", error);
                document.getElementById("output").innerText = "Erro ao processar os dados: " + error.message;
            }
        }
    </script>
</body>
</html>
