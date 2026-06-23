# Gerenciador de Inventário de Telefones VoIP

Aplicação CLI para gerenciar planilha de inventário de telefones VoIP.

## Requisitos

- Python 3.8+
- openpyxl

## Instalação

```bash
cd voip_cli
pip install -r requirements.txt
```

## Uso

```bash
python voip_cli/run.py list                    # Listar todos
python voip_cli/run.py list --status ONLINE    # Listar online
python voip_cli/run.py list --setor HEMOCENTRO # Listar por setor
python voip_cli/run.py search 3027-4407        # Buscar por IP/setor/telefone/MAC/serial
python voip_cli/run.py show 10.75.16.67        # Detalhes do telefone
python voip_cli/run.py add 10.75.17.1 "TI" 3027-1000 --modelo P11G --mac 00:11:22:33:44:55
python voip_cli/run.py edit 10.75.16.67 --setor "TI" --status OFFLINE
python voip_cli/run.py remove 10.75.16.67      # Remover telefone
python voip_cli/run.py ping                    # Testar conectividade (ping)
python voip_cli/run.py stats                   # Estatísticas do inventário
python voip_cli/run.py report --output relatorio.csv  # Exportar CSV
```

## Comandos

| Comando  | Descrição                          |
|----------|------------------------------------|
| list     | Listar telefones com filtros       |
| search   | Busca por termo em múltiplos campos|
| show     | Mostrar detalhes de um telefone    |
| add      | Adicionar novo telefone            |
| edit     | Editar telefone existente          |
| remove   | Remover telefone                   |
| ping     | Testar conectividade ICMP          |
| stats    | Exibir estatísticas                |
| report   | Exportar relatório CSV             |
