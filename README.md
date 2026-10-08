# Radar de Licitações 3D (Slim 3D)

Varre o PNCP todo dia útil atrás de editais com propostas abertas que envolvam
impressoras 3D, filamentos, resinas, scanners 3D e manufatura aditiva.

## O que ele entrega
Uma planilha por dia (`saida/radar_3d_AAAA-MM-DD.xlsx`) com:
- **Relevância**: Alta (3D no objeto), Média (item 3D escondido num lote), A confirmar
- **Novo?**: SIM para o que apareceu desde a última execução
- Data de encerramento das propostas, órgão, UF, itens 3D e valor estimado
- Link direto para o edital no PNCP
- Colunas "Status Slim" e "Observações" para o time comercial

Opcional: e-mail diário com os novos editais e a planilha anexa.

## Rodar no computador
```
pip install requests openpyxl
python radar_licitacoes_3d.py
```

## Rodar sozinho na nuvem (grátis, GitHub Actions)
1. Crie um repositório privado no GitHub e suba estes arquivos (incluindo a pasta `.github`).
2. Em Settings > Secrets and variables > Actions, cadastre:
   `SMTP_HOST` (ex.: smtp.gmail.com), `SMTP_PORT` (587), `SMTP_USER`,
   `SMTP_PASS` (senha de app do Gmail), `EMAIL_PARA`.
3. Pronto: roda de segunda a sexta às 7h. Também dá para rodar na aba Actions > Run workflow.

## Ajustes rápidos
- Termos de busca: lista `TERMOS_BUSCA` no topo do script
- O que conta como item 3D: `REGEX_3D`
- UFs em destaque: `UFS_PRIORITARIAS`

## Painel para o HubSpot
`painel_pregoes.html` é o painel visual com filtros. Cole o conteúdo inteiro num módulo
de HTML personalizado do HubSpot. Ele lê `radar.json` deste repositório
(https://raw.githubusercontent.com/jwhendel/radar3d/main/radar.json), atualizado todo dia
pelo agendamento. Se a leitura falhar, mostra o último retrato embutido.

`painel_template.html` é o modelo usado para gerar o painel (o marcador `__DADOS__`
recebe o retrato embutido).
