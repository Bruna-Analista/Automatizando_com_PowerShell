# ⚙️ Automação de Planilhas Excel com PowerShell (Renomeia planilhas + Exclui 3 primeiras linhas)

Scripts práticos de automação desenvolvidos em **PowerShell**  para otimizar tarefas do dia a dia em **análise de dados** e **controle de planilhas corporativas**.

### 🧹 1. Excluir as 3 primeiras linhas de várias planilhas
Remove automaticamente as três primeiras linhas de todos os arquivos `.xlsx` em uma pasta — sem abrir o Excel.

**🔹 Uso:**
1. Coloque todas as planilhas em uma pasta.
2. Ajuste o caminho da variável `$pasta` no script.
3. Execute o PowerShell como administrador e rode o comando:

## 💻 Script PowerShell

```powershell
# Caminho da pasta onde estão as planilhas
$pasta = "C:\Users\Bruna\Documents\planilhas_hospedagem"

# Cria uma instância invisível do Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

# Pega todos os arquivos .xlsx da pasta
$arquivos = Get-ChildItem -Path $pasta -Filter "*.xlsx"

foreach ($arquivo in $arquivos) {
    $caminhoCompleto = Join-Path $pasta $arquivo.Name
    Write-Host "Processando $($arquivo.Name)..."

    $workbook = $excel.Workbooks.Open($caminhoCompleto)
    $sheet = $workbook.Sheets.Item(1)

    # Exclui as 3 primeiras linhas
    $sheet.Rows("1:3").Delete()

    # Salva e fecha
    $workbook.Save()
    $workbook.Close()
}

$excel.Quit()

Write-Host "✅ Todas as planilhas foram atualizadas com sucesso!"

```
---

## ✏️Renomeando-varias-planilhas-ao-mesmo-tempo
Este repositório contém um script PowerShell simples e eficiente que renomeia diversos arquivos `.xlsx` de uma pasta, adicionando um sufixo personalizado no final do nome.

**🧠 Cenário de uso**

**Você tem vários arquivos como:**

- relatorio_vendas.xlsx
- extrato_funcionarios.xlsx
- alelo_gastos.xlsx


**E quer que eles fiquem assim:**


- relatorio_vendas_jun25.xlsx
- extrato_funcionarios_jun25.xlsx
- alelo_gastos_jun25.xlsx

**Abra o PowerShell (Windows+s e procure por powershell).**

Navegue até a pasta onde estão suas planilhas:
→ importante colocar o cd e aspas antes de inserir o caminho

cd "C:\caminho\da\pasta"

Dê enter.


**Depois copie e cole o Script abaixo:**


## 💻 Script PowerShell

```powershell
Get-ChildItem -Filter *.xlsx | ForEach-Object {
    $nomeOriginal = $_.BaseName
    $extensao = $_.Extension
    $novoNome = $nomeOriginal + "_abr2026" + $extensao
    Rename-Item -Path $_.FullName -NewName $novoNome
} 
```
---
## ✏️✅ Comando para COPIAR arquivos

```powershell
Copy-Item "C:\Users\bruna.zordenoni\Desktop\Bruna\Projects DEV Bruna\POWER BI\Desafio Junior\Portfolio Alelo - prestacao de contas\2026\*" "S:\BI\Custos\Alelo - prestacao de contas\2025" -Recurse
```
*O que faz:*

- Copy-Item = copiar
- \* = todos os arquivos da pasta
- -Recurse = inclui subpastas também

---
## ✅ Se quiser copiar apenas arquivos Excel
```powershell
Copy-Item "C:\Users\bruna.zordenoni\Desktop\Bruna\Projects DEV Bruna\POWER BI\Desafio Junior\Portfolio Alelo - prestacao de contas\2026\*.xlsx" "S:\BI\Custos\Alelo - prestacao de contas\2025"
```

## ✅ Se quiser MOVER (tirar da origem e colocar no destino)

```powershell
Move-Item "C:\Origem\*.xlsx" "D:\Destino\"
```


