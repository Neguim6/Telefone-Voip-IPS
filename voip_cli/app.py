import argparse
import sys
import subprocess
import platform
import getpass
from pathlib import Path
from .database import VoipDatabase, COLUMNS

EXCEL_PATH = Path(__file__).parent.parent / "Telefones VOIPS atualizados 15 04 2026.xlsx"
SENHA = "258456"


def require_auth(password_arg: str | None = None) -> bool:
    if password_arg == SENHA:
        return True
    if password_arg is None:
        senha = getpass.getpass("Senha: ")
        if senha == SENHA:
            return True
    print("Senha incorreta! Acesso negado.")
    return False


def get_db():
    try:
        return VoipDatabase(str(EXCEL_PATH))
    except FileNotFoundError:
        print(f"Erro: Arquivo Excel não encontrado em {EXCEL_PATH}")
        sys.exit(1)


def cmd_list(args):
    db = get_db()
    data = db.list_all()

    if args.setor:
        data = [r for r in data if args.setor.lower() in r.get("Setor", "").lower()]
    if args.status:
        data = [r for r in data if r.get("Status", "").upper() == args.status.upper()]
    if args.modelo:
        data = [r for r in data if args.modelo.lower() in r.get("Nome do Produto", "").lower()]

    if not data:
        print("Nenhum registro encontrado.")
        return

    print(f"{'IP':<16} {'Setor':<40} {'Telefone':<14} {'Status':<10} {'Modelo':<8}")
    print("-" * 90)
    for r in data:
        setor = (r.get("Setor") or "")[:38]
        print(f"{r.get('IP', ''):<16} {setor:<40} {r.get('Telefone', ''):<14} {r.get('Status', ''):<10} {r.get('Nome do Produto', ''):<8}")
    print(f"\nTotal: {len(data)} registros")


def cmd_search(args):
    db = get_db()
    data = db.search(IP=args.termo, Setor=args.termo, Telefone=args.termo, MAC=args.termo, Serial=args.termo)

    if not data:
        print("Nenhum resultado encontrado.")
        return

    for r in data:
        print(f"\nIP:       {r.get('IP', '')}")
        print(f"Setor:    {r.get('Setor', '')}")
        print(f"Telefone: {r.get('Telefone', '')}")
        print(f"Status:   {r.get('Status', '')}")
        print(f"Modelo:   {r.get('Nome do Produto', '')}")
        print(f"MAC:      {r.get('MAC', '')}")
        print(f"Serial:   {r.get('Serial', '')}")
        print(f"Firmware: {r.get('Firmware', '')}")


def cmd_show(args):
    db = get_db()
    r = db.get_by_ip(args.ip)
    if not r:
        print(f"Telefone com IP {args.ip} não encontrado.")
        return

    print(f"IP:       {r.get('IP', '')}")
    print(f"Setor:    {r.get('Setor', '')}")
    print(f"Telefone: {r.get('Telefone', '')}")
    print(f"Status:   {r.get('Status', '')}")
    print(f"Modelo:   {r.get('Nome do Produto', '')}")
    print(f"MAC:      {r.get('MAC', '')}")
    print(f"Hardware: {r.get('Hardware', '')}")
    print(f"Firmware: {r.get('Firmware', '')}")
    print(f"Serial:   {r.get('Serial', '')}")


def cmd_add(args):
    if not require_auth(args.password):
        return
    db = get_db()
    record = {
        "IP": args.ip,
        "Setor": args.setor,
        "Telefone": args.telefone,
        "Status": args.status or "ONLINE",
        "Nome do Produto": args.modelo or "",
        "MAC": args.mac or "",
        "Hardware": args.hardware or "",
        "Firmware": args.firmware or "",
        "Serial": args.serial or "",
    }
    db.add(record)
    print(f"Telefone {args.ip} adicionado com sucesso!")


def cmd_edit(args):
    if not require_auth(args.password):
        return
    db = get_db()
    updates = {}
    if args.setor:
        updates["Setor"] = args.setor
    if args.telefone:
        updates["Telefone"] = args.telefone
    if args.status:
        updates["Status"] = args.status
    if args.modelo:
        updates["Nome do Produto"] = args.modelo
    if args.mac:
        updates["MAC"] = args.mac
    if args.serial:
        updates["Serial"] = args.serial
    if args.firmware:
        updates["Firmware"] = args.firmware

    if db.update(args.ip, updates):
        print(f"Telefone {args.ip} atualizado com sucesso!")
    else:
        print(f"Telefone com IP {args.ip} não encontrado.")


def cmd_remove(args):
    if not require_auth(args.password):
        return
    db = get_db()
    if args.yes or input(f"Remover telefone {args.ip}? (s/N): ").lower() == "s":
        if db.delete(args.ip):
            print(f"Telefone {args.ip} removido com sucesso!")
        else:
            print(f"Telefone com IP {args.ip} não encontrado.")


def cmd_ping(args):
    db = get_db()
    param = "-n" if platform.system().lower() == "windows" else "-c"

    data = db.list_all()
    if args.setor:
        data = [r for r in data if args.setor.lower() in r.get("Setor", "").lower()]

    if not data:
        print("Nenhum registro encontrado.")
        return

    online_count = 0
    offline_count = 0
    for r in data:
        ip = r.get("IP", "")
        if not ip:
            continue
        try:
            result = subprocess.run(
                ["ping", param, "1", ip],
                capture_output=True,
                text=True,
                timeout=5,
            )
            if result.returncode == 0:
                print(f"  ONLINE   {ip}  -  {r.get('Setor', '')}")
                online_count += 1
            else:
                print(f"  OFFLINE  {ip}  -  {r.get('Setor', '')}")
                offline_count += 1
        except subprocess.TimeoutExpired:
            print(f"  TIMEOUT  {ip}  -  {r.get('Setor', '')}")
            offline_count += 1

    print(f"\nTotal: {len(data)} | Online: {online_count} | Offline: {offline_count}")

    if not args.no_update:
        if not require_auth(args.password):
            return
        for r in data:
            ip = r.get("IP", "")
            if not ip:
                continue
            result = subprocess.run(
                ["ping", param, "1", ip],
                capture_output=True,
                text=True,
                timeout=5,
            )
            new_status = "ONLINE" if result.returncode == 0 else "OFFLINE"
            db.update(ip, {"Status": new_status})
        print("Status atualizado na planilha.")


def cmd_stats(args):
    db = get_db()
    s = db.stats()
    print(f"\n=== Estatísticas do Inventário VoIP ===\n")
    print(f"Total de telefones:    {s['total']}")
    print(f"Online:                {s['online']}")
    print(f"Offline:               {s['offline']}")
    print(f"Setores distintos:     {s['setores_unicos']}")
    print(f"\nModelos:")
    for modelo, qtd in sorted(s["modelos"].items(), key=lambda x: -x[1]):
        print(f"  {modelo}: {qtd}")


def cmd_report(args):
    db = get_db()
    data = db.list_all()

    if args.setor:
        data = [r for r in data if args.setor.lower() in r.get("Setor", "").lower()]

    if args.output:
        import csv
        with open(args.output, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=COLUMNS)
            writer.writeheader()
            writer.writerows(data)
        print(f"Relatório exportado para {args.output}")
    else:
        print(f"{'IP':<16} {'Setor':<40} {'Telefone':<14} {'Status':<10} {'MAC':<20}")
        print("-" * 100)
        for r in data:
            setor = (r.get("Setor") or "")[:38]
            print(f"{r.get('IP', ''):<16} {setor:<40} {r.get('Telefone', ''):<14} {r.get('Status', ''):<10} {r.get('MAC', ''):<20}")
        print(f"\nTotal: {len(data)} registros")


def main():
    parser = argparse.ArgumentParser(
        description="Gerenciador de Inventário de Telefones VoIP",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Comandos disponíveis:
  list     Listar todos os telefones
  search   Buscar telefones por termo
  show     Mostrar detalhes de um telefone
  add      Adicionar novo telefone
  edit     Editar um telefone existente
  remove   Remover um telefone
  ping     Testar conectividade dos telefones
  stats    Exibir estatísticas do inventário
  report   Exportar relatório (CSV ou tela)

Exemplos:
  voip list
  voip list --status ONLINE
  voip list --setor "HEMOCENTRO"
  voip search 3027-4407
  voip show 10.75.16.67
  voip add 10.75.17.1 "TI" 3027-1000 --modelo P11G --mac 00:11:22:33:44:55
  voip edit 10.75.16.67 --setor "TI" --status OFFLINE
  voip ping --setor HEMOCENTRO
  voip stats
  voip report --output relatorio.csv
        """
    )
    subparsers = parser.add_subparsers(dest="command")

    # list
    p_list = subparsers.add_parser("list", help="Listar telefones")
    p_list.add_argument("--setor", "-s", help="Filtrar por setor")
    p_list.add_argument("--status", "-st", help="Filtrar por status (ONLINE/OFFLINE)")
    p_list.add_argument("--modelo", "-m", help="Filtrar por modelo")
    p_list.set_defaults(func=cmd_list)

    # search
    p_search = subparsers.add_parser("search", help="Buscar telefones")
    p_search.add_argument("termo", help="Termo para busca (IP, setor, telefone, MAC, serial)")
    p_search.set_defaults(func=cmd_search)

    # show
    p_show = subparsers.add_parser("show", help="Mostrar detalhes de um telefone")
    p_show.add_argument("ip", help="IP do telefone")
    p_show.set_defaults(func=cmd_show)

    # add
    p_add = subparsers.add_parser("add", help="Adicionar novo telefone")
    p_add.add_argument("ip", help="Endereço IP")
    p_add.add_argument("setor", help="Setor/Departamento")
    p_add.add_argument("telefone", help="Número de telefone")
    p_add.add_argument("--status", help="Status (padrão: ONLINE)")
    p_add.add_argument("--modelo", help="Nome do produto/modelo")
    p_add.add_argument("--mac", help="Endereço MAC")
    p_add.add_argument("--hardware", help="Versão do hardware")
    p_add.add_argument("--firmware", help="Versão do firmware")
    p_add.add_argument("--serial", help="Número de série")
    p_add.add_argument("--password", "-p", help="Senha para alteração")
    p_add.set_defaults(func=cmd_add)

    # edit
    p_edit = subparsers.add_parser("edit", help="Editar um telefone")
    p_edit.add_argument("ip", help="IP do telefone a editar")
    p_edit.add_argument("--setor", help="Novo setor")
    p_edit.add_argument("--telefone", help="Novo número")
    p_edit.add_argument("--status", help="Novo status")
    p_edit.add_argument("--modelo", help="Novo modelo")
    p_edit.add_argument("--mac", help="Novo MAC")
    p_edit.add_argument("--serial", help="Novo serial")
    p_edit.add_argument("--firmware", help="Novo firmware")
    p_edit.add_argument("--password", "-p", help="Senha para alteração")
    p_edit.set_defaults(func=cmd_edit)

    # remove
    p_remove = subparsers.add_parser("remove", help="Remover um telefone")
    p_remove.add_argument("ip", help="IP do telefone a remover")
    p_remove.add_argument("--yes", "-y", action="store_true", help="Confirmação automática")
    p_remove.add_argument("--password", "-p", help="Senha para alteração")
    p_remove.set_defaults(func=cmd_remove)

    # ping
    p_ping = subparsers.add_parser("ping", help="Testar conectividade dos telefones")
    p_ping.add_argument("--setor", "-s", help="Filtrar por setor")
    p_ping.add_argument("--no-update", action="store_true", help="Não atualizar status na planilha")
    p_ping.add_argument("--password", "-p", help="Senha para alteração (necessária se for atualizar status)")
    p_ping.set_defaults(func=cmd_ping)

    # stats
    p_stats = subparsers.add_parser("stats", help="Exibir estatísticas")
    p_stats.set_defaults(func=cmd_stats)

    # report
    p_report = subparsers.add_parser("report", help="Exportar relatório")
    p_report.add_argument("--setor", "-s", help="Filtrar por setor")
    p_report.add_argument("--output", "-o", help="Arquivo de saída CSV")
    p_report.set_defaults(func=cmd_report)

    args = parser.parse_args()
    if args.command is None:
        parser.print_help()
        sys.exit(1)

    args.func(args)


if __name__ == "__main__":
    main()
