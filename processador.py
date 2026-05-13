import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
import re

NS = {'nfse': 'http://www.sped.fazenda.gov.br/nfse'}

MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
}


def identificar_retencoes(xml_path):
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        retido = False

        tp_ret_issqn = root.find('.//nfse:tpRetISSQN', NS)
        if tp_ret_issqn is not None and tp_ret_issqn.text == '1':
            retido = True

        tp_ret_piscofins = root.find('.//nfse:tpRetPisCofins', NS)
        if tp_ret_piscofins is not None and tp_ret_piscofins.text == '1':
            retido = True

        for tag in ('nfse:vRetIRRF', 'nfse:vRetCP', 'nfse:vRetCSLL'):
            elem = root.find(f'.//{tag}', NS)
            if elem is not None:
                try:
                    if float(elem.text) > 0:
                        retido = True
                except (ValueError, TypeError):
                    pass

        return retido
    except ET.ParseError:
        return None


def extrair_mes_ano_do_nome(nome):
    match = re.search(r'(\d{2})-(\d{2})-(\d{4})', nome)
    if match:
        return int(match.group(2)), int(match.group(3))
    return None, None


def encontrar_pasta_mes(pasta_base, mes, ano):
    nome_mes = MESES_PT.get(mes, '').upper()
    if not os.path.exists(pasta_base):
        return None
    for item in os.listdir(pasta_base):
        if nome_mes in item.upper():
            return os.path.join(pasta_base, item)
    return None


def criar_estrutura_pastas(pasta_servidor, nome_empresa_servidor, mes, ano, tipo):
    caminho_base = os.path.join(
        pasta_servidor, nome_empresa_servidor,
        'Fiscal', 'Fechamento Fiscal', str(ano)
    )
    pasta_mes = encontrar_pasta_mes(caminho_base, mes, ano)
    caminho_mes = pasta_mes if pasta_mes else os.path.join(caminho_base, MESES_PT[mes])

    if tipo == 'Tomadas':
        caminho_tomadas = os.path.join(caminho_mes, 'Tomadas')
        retidas     = os.path.join(caminho_tomadas, 'Retidas')
        nao_retidas = os.path.join(caminho_tomadas, 'Não Retidas')
        canceladas  = os.path.join(caminho_tomadas, 'Canceladas')
        os.makedirs(retidas, exist_ok=True)
        os.makedirs(nao_retidas, exist_ok=True)
        os.makedirs(canceladas, exist_ok=True)
        return retidas, nao_retidas, canceladas
    else:
        emitidas = os.path.join(caminho_mes, 'Emitidas')
        os.makedirs(emitidas, exist_ok=True)
        return emitidas, None, None


def carregar_mapeamento(caminho_config, callback_log=None):
    def log(msg):
        if callback_log:
            callback_log(msg)

    if not caminho_config or not os.path.exists(caminho_config):
        return {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(caminho_config)
        ws = wb.active
        return {
            str(row[0]).strip().upper(): str(row[1]).strip()
            for row in ws.iter_rows(min_row=2, values_only=True)
            if row[0] and row[1]
        }
    except Exception as e:
        log(f"❌ Erro ao carregar mapeamento: {e}")
        return {}


def _mover_par(xml_path, destino_dir, log):
    """Move o XML e o PDF de mesmo nome para destino_dir. Retorna True se moveu."""
    arq = os.path.basename(xml_path)
    dest_xml = os.path.join(destino_dir, arq)
    if os.path.exists(dest_xml):
        log(f"  ⚠️ Já existe, ignorando: {arq}")
        os.remove(xml_path)
        pdf_src = os.path.splitext(xml_path)[0] + '.pdf'
        if os.path.exists(pdf_src):
            os.remove(pdf_src)
        return False
    shutil.move(xml_path, dest_xml)
    pdf_src = os.path.splitext(xml_path)[0] + '.pdf'
    if os.path.exists(pdf_src):
        shutil.move(pdf_src, os.path.join(destino_dir, os.path.basename(pdf_src)))
    return True


def _encontrar_pastas_periodo(pasta_empresa):
    """
    Localiza subpastas de período (ex: '01-03-2026 a 31-03-2026 Recebidas').
    Suporta nesting duplo (EMPRESA/EMPRESA/período).
    Retorna lista de (caminho_completo, nome_pasta).
    """
    resultado = []

    def _buscar(pasta, profundidade):
        if profundidade > 2:
            return
        try:
            for item in os.listdir(pasta):
                caminho = os.path.join(pasta, item)
                if not os.path.isdir(caminho):
                    continue
                mes, ano = extrair_mes_ano_do_nome(item)
                if mes and ano:
                    resultado.append((caminho, item))
                else:
                    _buscar(caminho, profundidade + 1)
        except PermissionError:
            pass

    _buscar(pasta_empresa, 0)
    return resultado


def processar_empresa(pasta_empresa, nome_empresa_portal, pasta_servidor,
                      mapeamento, callback_log=None):
    stats = {
        'empresa': nome_empresa_portal,
        'retidas': 0,
        'nao_retidas': 0,
        'emitidas': 0,
        'canceladas_movidas': 0,
        'erros': 0,
    }

    nome_upper = nome_empresa_portal.strip().upper()
    nome_servidor = mapeamento.get(nome_upper, nome_empresa_portal)

    def log(msg):
        if callback_log:
            callback_log(msg)

    log(f"📂 Processando: {nome_empresa_portal}")

    pastas_periodo = _encontrar_pastas_periodo(pasta_empresa)
    if not pastas_periodo:
        log(f"  ⚠️ Nenhuma pasta de período encontrada.")
        return stats

    for caminho_periodo, nome_periodo in pastas_periodo:
        nome_upper_p = nome_periodo.upper()
        mes, ano = extrair_mes_ano_do_nome(nome_periodo)

        if 'EMITID' in nome_upper_p:
            tipo = 'Emitidas'
        elif 'RECEBID' in nome_upper_p or 'TOMAD' in nome_upper_p:
            tipo = 'Tomadas'
        else:
            log(f"  ⚠️ Tipo não identificado: {nome_periodo}")
            continue

        dest_a, dest_b, dest_canceladas = criar_estrutura_pastas(
            pasta_servidor, nome_servidor, mes, ano, tipo
        )

        cnt_a = cnt_b = cnt_c = cnt_err = 0

        for raiz, dirs, arquivos in os.walk(caminho_periodo):
            pasta_nome = os.path.basename(raiz).upper()

            if 'CANCELAD' in pasta_nome and raiz != caminho_periodo:
                # Move canceladas para pasta Canceladas no servidor
                if tipo == 'Tomadas':
                    for arq in arquivos:
                        if arq.lower().endswith('.xml'):
                            xml_path = os.path.join(raiz, arq)
                            if _mover_par(xml_path, dest_canceladas, log):
                                cnt_c += 1
                                stats['canceladas_movidas'] += 1
                dirs[:] = []
                continue

            for arq in arquivos:
                if not arq.lower().endswith('.xml'):
                    continue
                xml_path = os.path.join(raiz, arq)

                if tipo == 'Emitidas':
                    if _mover_par(xml_path, dest_a, log):
                        cnt_a += 1
                        stats['emitidas'] += 1
                else:
                    retido = identificar_retencoes(xml_path)
                    if retido is None:
                        log(f"  ❌ XML inválido: {arq}")
                        cnt_err += 1
                        stats['erros'] += 1
                        continue
                    if retido:
                        if _mover_par(xml_path, dest_a, log):
                            cnt_a += 1
                            stats['retidas'] += 1
                    else:
                        if _mover_par(xml_path, dest_b, log):
                            cnt_b += 1
                            stats['nao_retidas'] += 1

        if tipo == 'Emitidas':
            log(f"  ✅ {nome_periodo}: {cnt_a} emitida(s)")
        else:
            log(f"  ✅ {nome_periodo}: Retidas={cnt_a} | Não Retidas={cnt_b} | Canceladas={cnt_c}")

    return stats


def executar_automacao(pasta_origem, pasta_servidor, caminho_config,
                       callback_log=None, callback_progresso=None):
    def log(msg):
        if callback_log:
            callback_log(msg)

    if not os.path.exists(pasta_origem):
        log("❌ Pasta de origem não encontrada.")
        return
    if not os.path.exists(pasta_servidor):
        log("❌ Pasta do servidor não encontrada.")
        return

    mapeamento = carregar_mapeamento(caminho_config, callback_log)
    if not mapeamento:
        log("⚠️ Mapeamento não carregado. Usando nomes do portal.")

    # Coleta empresas: diretórios e ZIPs na pasta de origem
    empresas = []   # lista de (nome_empresa, caminho_pasta, pasta_temp_para_limpar)
    pastas_temp = []

    for item in os.listdir(pasta_origem):
        caminho = os.path.join(pasta_origem, item)

        if os.path.isdir(caminho):
            nome = item
            empresas.append((nome, caminho, None))

        elif item.lower().endswith('.zip'):
            nome = os.path.splitext(item)[0]
            pasta_temp = os.path.join(pasta_origem, f"_temp_zip_{nome}")
            try:
                with zipfile.ZipFile(caminho, 'r') as z:
                    z.extractall(pasta_temp)
                # Após extração, a empresa pode estar um nível abaixo (pasta com mesmo nome)
                sub = os.path.join(pasta_temp, nome)
                pasta_empresa = sub if os.path.isdir(sub) else pasta_temp
                empresas.append((nome, pasta_empresa, pasta_temp))
                pastas_temp.append(pasta_temp)
            except zipfile.BadZipFile:
                log(f"❌ Não foi possível abrir ZIP: {item}")

    if not empresas:
        log("⚠️ Nenhuma empresa encontrada na pasta de origem.")
        return

    log(f"🚀 Iniciando processamento de {len(empresas)} empresa(s)...\n")

    stats_total = {
        'retidas': 0, 'nao_retidas': 0,
        'emitidas': 0, 'canceladas_movidas': 0, 'erros': 0
    }

    for i, (nome, pasta, _) in enumerate(empresas):
        stats = processar_empresa(
            pasta, nome, pasta_servidor,
            mapeamento, callback_log
        )
        for chave in stats_total:
            stats_total[chave] += stats.get(chave, 0)

        if callback_progresso:
            callback_progresso(int((i + 1) / len(empresas) * 100))
        log("")

    # Remove pastas temporárias de ZIPs
    for pt in pastas_temp:
        shutil.rmtree(pt, ignore_errors=True)

    log("=" * 50)
    log("✅ PROCESSAMENTO CONCLUÍDO")
    log(f"   Retidas:          {stats_total['retidas']}")
    log(f"   Não Retidas:      {stats_total['nao_retidas']}")
    log(f"   Emitidas:         {stats_total['emitidas']}")
    log(f"   Canceladas:       {stats_total['canceladas_movidas']}")
    log(f"   Erros:            {stats_total['erros']}")
    log("=" * 50)
