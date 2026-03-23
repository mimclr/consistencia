import os
import re
import tempfile
from dataclasses import dataclass, asdict
from typing import List, Dict, Any

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

APP_TITLE = 'Validador MIMC'


@dataclass
class Inconsistencia:
    aba: str
    linha: int
    regra: str
    severidade: str
    descricao: str
    valor_encontrado: str = ''
    valor_esperado: str = ''


def normalizar_colunas(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def procurar_coluna(df: pd.DataFrame, candidatos: List[str]) -> str | None:
    cols_norm = {re.sub(r'\s+', ' ', str(c).strip().upper()): c for c in df.columns}
    for cand in candidatos:
        key = re.sub(r'\s+', ' ', cand.strip().upper())
        if key in cols_norm:
            return cols_norm[key]
    for c in df.columns:
        cu = str(c).upper()
        for cand in candidatos:
            if cand.upper() in cu:
                return c
    return None


def eh_sim(v: Any) -> bool:
    return str(v).strip().upper() in {'SIM', 'S', 'TRUE', '1'}


def texto(v: Any) -> str:
    if pd.isna(v):
        return ''
    return str(v).strip()


def numero(v: Any):
    if pd.isna(v) or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    s = s.replace('.', '').replace(',', '.') if s.count(',') == 1 and s.count('.') >= 1 else s.replace(',', '.')
    try:
        return float(s)
    except Exception:
        return None


def detectar_mao_de_obra(valor: str) -> bool:
    v = texto(valor).upper()
    termos = [
        'MAO DE OBRA', 'MÃO DE OBRA', 'DIARIA', 'DIÁRIA', 'DIARISTA', 'SERVICO', 'SERVIÇO',
        'OPERADOR', 'TRABALHADOR', 'ROÇADA MANUAL', 'APLICACAO MANUAL', 'APLICAÇÃO MANUAL',
        'COLHEITA MANUAL'
    ]
    return any(t in v for t in termos)


def analisar_planilha(caminho: str) -> List[Inconsistencia]:
    inconsistencias: List[Inconsistencia] = []
    xls = pd.ExcelFile(caminho, engine='openpyxl')
    abas = set(xls.sheet_names)

    obrigatorias = {'INVENTARIO', 'PRODUCAO', 'VENDAS', 'DESPESAS', 'TALHAO'}
    faltantes = obrigatorias - abas
    for aba in sorted(faltantes):
        inconsistencias.append(Inconsistencia(
            aba='ESTRUTURA',
            linha=0,
            regra='EST-001',
            severidade='CRÍTICO',
            descricao=f'Aba obrigatória ausente: {aba}'
        ))

    talhoes_producao = []
    talhoes_todos = []
    if 'TALHAO' in abas:
        df_talhao = normalizar_colunas(pd.read_excel(xls, 'TALHAO'))
        col_talhao = procurar_coluna(df_talhao, ['TALHÃO', 'TALHAO'])
        col_estagio = procurar_coluna(df_talhao, ['ESTÁGIO', 'ESTAGIO'])
        if col_talhao:
            talhoes_todos = [texto(v) for v in df_talhao[col_talhao].tolist() if texto(v)]
        if col_talhao and col_estagio:
            talhoes_producao = [
                texto(row[col_talhao])
                for _, row in df_talhao.iterrows()
                if texto(row[col_estagio]).upper() == 'PRODUÇÃO' or texto(row[col_estagio]).upper() == 'PRODUCAO'
            ]

    if 'INVENTARIO' in abas:
        df = normalizar_colunas(pd.read_excel(xls, 'INVENTARIO'))
        col_item_novo = procurar_coluna(df, ['VALOR DO ITEM NOVO (R$)', 'VALOR DO ITEM NOVO'])
        col_valor_pago = procurar_coluna(df, ['VALOR PAGO (R$)', 'VALOR PAGO'])
        col_fab = procurar_coluna(df, ['DATA DE FABRICAÇÃO', 'DATA DE FABRICACAO'])
        col_aq = procurar_coluna(df, ['DATA DE AQUISIÇÃO', 'DATA DE AQUISICAO'])

        for i, row in df.iterrows():
            for col in [col_item_novo, col_valor_pago]:
                if col:
                    val = numero(row[col])
                    if val is not None and (val < 100 or val > 500000):
                        inconsistencias.append(Inconsistencia(
                            aba='INVENTARIO', linha=i + 2, regra='INV-001', severidade='MÉDIO',
                            descricao=f'Valor fora da faixa na coluna {col}',
                            valor_encontrado=texto(row[col]), valor_esperado='Entre 100 e 500000'
                        ))
            if col_fab and col_aq and pd.notna(row[col_fab]) and pd.notna(row[col_aq]):
                try:
                    dfab = pd.to_datetime(row[col_fab], dayfirst=True, errors='coerce')
                    daq = pd.to_datetime(row[col_aq], dayfirst=True, errors='coerce')
                    if pd.notna(dfab) and pd.notna(daq) and dfab > daq:
                        inconsistencias.append(Inconsistencia(
                            aba='INVENTARIO', linha=i + 2, regra='INV-002', severidade='ALTO',
                            descricao='Data de fabricação maior que data de aquisição',
                            valor_encontrado=f'{dfab.date()} > {daq.date()}',
                            valor_esperado='Fabricação menor ou igual à aquisição'
                        ))
                except Exception:
                    pass

    if 'VENDAS' in abas:
        df = normalizar_colunas(pd.read_excel(xls, 'VENDAS'))
        col_preco = procurar_coluna(df, ['PREÇO DE VENDA (R$/SC)', 'PRECO DE VENDA (R$/SC)', 'PREÇO DE VENDA', 'PRECO DE VENDA'])
        for i, row in df.iterrows():
            if col_preco:
                val = numero(row[col_preco])
                if val is not None and val > 100:
                    inconsistencias.append(Inconsistencia(
                        aba='VENDAS', linha=i + 2, regra='VEN-001', severidade='ALTO',
                        descricao='Preço de venda acima de 100',
                        valor_encontrado=texto(row[col_preco]), valor_esperado='Até 100'
                    ))

    if 'PRODUCAO' in abas:
        df = normalizar_colunas(pd.read_excel(xls, 'PRODUCAO'))
        col_rateio = procurar_coluna(df, ['RATEIO'])
        col_talhao = procurar_coluna(df, ['TALHÃO', 'TALHAO'])
        col_mes = procurar_coluna(df, ['MÊS', 'MES'])
        col_safra = procurar_coluna(df, ['SAFRA'])
        col_producao = procurar_coluna(df, ['PRODUÇÃO TOTAL', 'PRODUCAO TOTAL'])
        if col_rateio and col_talhao:
            rateados = df[df[col_rateio].apply(eh_sim)].copy()
            if not rateados.empty:
                group_cols = [c for c in [col_mes, col_safra] if c]
                if not group_cols:
                    rateados['_GRUPO_'] = 'UNICO'
                    group_cols = ['_GRUPO_']
                for _, g in rateados.groupby(group_cols, dropna=False):
                    presentes = {texto(v) for v in g[col_talhao].tolist() if texto(v)}
                    faltantes = sorted(set(talhoes_producao) - presentes)
                    if faltantes:
                        inconsistencias.append(Inconsistencia(
                            aba='PRODUCAO', linha=int(g.index.min()) + 2, regra='PRO-001', severidade='CRÍTICO',
                            descricao='Rateio sem todos os talhões em produção',
                            valor_encontrado=', '.join(faltantes), valor_esperado='Todos os talhões em produção'
                        ))
                    if col_producao:
                        valores = {numero(v) for v in g[col_producao].tolist() if numero(v) is not None}
                        if len(valores) > 1:
                            inconsistencias.append(Inconsistencia(
                                aba='PRODUCAO', linha=int(g.index.min()) + 2, regra='PRO-002', severidade='CRÍTICO',
                                descricao='Valores diferentes em lançamento com rateio',
                                valor_encontrado=', '.join(map(lambda x: str(x), sorted(valores))),
                                valor_esperado='Mesmo valor para todos os talhões do grupo'
                            ))

    if 'DESPESAS' in abas:
        df = normalizar_colunas(pd.read_excel(xls, 'DESPESAS'))
        col_rateio = procurar_coluna(df, ['RATEIO'])
        col_talhao = procurar_coluna(df, ['TALHÃO', 'TALHAO'])
        col_mes = procurar_coluna(df, ['MÊS', 'MES'])
        col_atividade = procurar_coluna(df, ['ATIVIDADE'])
        col_elemento = procurar_coluna(df, ['ELEMENTO', 'INSUMO', 'DESCRIÇÃO', 'DESCRICAO'])
        col_valor_total = procurar_coluna(df, ['VALOR TOTAL', 'VALOR'])
        atividades_mo = {
            'CONDUÇÃO DE LAVOURA', 'CONDUCAO DE LAVOURA', 'ADUBAÇÃO VIA SOLO', 'ADUBACAO VIA SOLO',
            'ADUBAÇÃO VIA FOLHA', 'ADUBACAO VIA FOLHA', 'CONTROLE DE PRAGAS E DOENÇAS',
            'CONTROLE DE PRAGAS E DOENCAS', 'CONTROLE DE PLANTAS DANINHAS', 'COLHEITA',
            'PÓS-COLHEITA', 'POS-COLHEITA', 'POS COLHEITA'
        }

        if col_atividade and col_mes:
            adm_mask = df[col_atividade].astype(str).str.upper().str.contains('ADMINISTRA', na=False)
            meses_adm = []
            for v in df.loc[adm_mask, col_mes].dropna().tolist():
                try:
                    dt = pd.to_datetime(v, dayfirst=True, errors='coerce')
                    if pd.notna(dt):
                        meses_adm.append(dt.to_period('M'))
                except Exception:
                    pass
            if meses_adm:
                ini = min(meses_adm)
                fim = max(meses_adm)
                atual = ini
                existentes = set(meses_adm)
                while atual <= fim:
                    if atual not in existentes:
                        inconsistencias.append(Inconsistencia(
                            aba='DESPESAS', linha=0, regra='DES-001', severidade='MÉDIO',
                            descricao='Falha de recorrência administrativa',
                            valor_encontrado=str(atual), valor_esperado='Mês com lançamento de administração'
                        ))
                    atual += 1

        if col_atividade and col_mes and col_elemento:
            base = df.copy()
            base['_ATV_'] = base[col_atividade].apply(lambda x: texto(x).upper())
            base['_MES_'] = base[col_mes].apply(texto)
            for (mes, atv), g in base.groupby(['_MES_', '_ATV_'], dropna=False):
                if atv in atividades_mo:
                    tem_insumo = any(not detectar_mao_de_obra(v) and texto(v) for v in g[col_elemento].tolist())
                    tem_mao = any(detectar_mao_de_obra(v) for v in g[col_elemento].tolist())
                    if tem_insumo and not tem_mao:
                        inconsistencias.append(Inconsistencia(
                            aba='DESPESAS', linha=int(g.index.min()) + 2, regra='DES-002', severidade='ALTO',
                            descricao=f'Atividade {atv} com lançamento sem mão de obra associada',
                            valor_encontrado=f'Mês {mes}', valor_esperado='Existência de mão de obra associada'
                        ))

        if col_atividade:
            for i, row in df.iterrows():
                atv = texto(row[col_atividade]).upper()
                if 'MANUTENÇÃO DE MÁQUINAS' in atv or 'MANUTENCAO DE MAQUINAS' in atv or 'MANUTENÇÃO DE MÁQUINAS E IMPLEMENTOS' in atv or 'MANUTENCAO DE MAQUINAS E IMPLEMENTOS' in atv:
                    if 'ADMINISTRA' in atv:
                        inconsistencias.append(Inconsistencia(
                            aba='DESPESAS', linha=i + 2, regra='DES-003', severidade='ALTO',
                            descricao='Manutenção de máquinas/equipamentos lançada como administração'
                        ))

        if col_rateio and col_talhao:
            rateados = df[df[col_rateio].apply(eh_sim)].copy()
            if not rateados.empty:
                group_cols = [c for c in [col_mes, col_atividade, col_elemento] if c]
                if not group_cols:
                    rateados['_GRUPO_'] = 'UNICO'
                    group_cols = ['_GRUPO_']
                for _, g in rateados.groupby(group_cols, dropna=False):
                    presentes = {texto(v) for v in g[col_talhao].tolist() if texto(v)}
                    faltantes = sorted(set(talhoes_todos) - presentes)
                    if faltantes:
                        inconsistencias.append(Inconsistencia(
                            aba='DESPESAS', linha=int(g.index.min()) + 2, regra='DES-004', severidade='CRÍTICO',
                            descricao='Rateio sem todos os talhões cadastrados',
                            valor_encontrado=', '.join(faltantes), valor_esperado='Todos os talhões cadastrados'
                        ))
                    if col_valor_total:
                        valores = {numero(v) for v in g[col_valor_total].tolist() if numero(v) is not None}
                        if len(valores) > 1:
                            inconsistencias.append(Inconsistencia(
                                aba='DESPESAS', linha=int(g.index.min()) + 2, regra='DES-005', severidade='CRÍTICO',
                                descricao='Valores diferentes em despesa com rateio',
                                valor_encontrado=', '.join(map(lambda x: str(x), sorted(valores))),
                                valor_esperado='Mesmo valor para todos os talhões do grupo'
                            ))

    return inconsistencias


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry('1050x620')
        self.caminho_arquivo = ''
        self.resultados: List[Inconsistencia] = []
        self._montar_ui()

    def _montar_ui(self):
        topo = tk.Frame(self.root)
        topo.pack(fill='x', padx=12, pady=12)

        tk.Label(topo, text=APP_TITLE, font=('Arial', 16, 'bold')).pack(anchor='w')
        tk.Label(topo, text='Validação de 1 planilha por vez').pack(anchor='w', pady=(0, 10))

        linha = tk.Frame(topo)
        linha.pack(fill='x')
        tk.Button(linha, text='Selecionar planilha', command=self.selecionar_arquivo, width=20).pack(side='left')
        tk.Button(linha, text='Analisar', command=self.analisar, width=15).pack(side='left', padx=8)
        tk.Button(linha, text='Exportar relatório', command=self.exportar, width=18).pack(side='left')

        self.lbl_arquivo = tk.Label(topo, text='Nenhum arquivo selecionado', anchor='w')
        self.lbl_arquivo.pack(fill='x', pady=(10, 0))

        resumo = tk.Frame(self.root)
        resumo.pack(fill='x', padx=12)
        self.lbl_resumo = tk.Label(resumo, text='Aguardando análise', anchor='w', justify='left')
        self.lbl_resumo.pack(fill='x')

        cols = ('aba', 'linha', 'regra', 'severidade', 'descricao', 'valor_encontrado', 'valor_esperado')
        tabela_frame = tk.Frame(self.root)
        tabela_frame.pack(fill='both', expand=True, padx=12, pady=12)
        self.tree = ttk.Treeview(tabela_frame, columns=cols, show='headings')
        headings = {
            'aba': 'Aba', 'linha': 'Linha', 'regra': 'Regra', 'severidade': 'Severidade',
            'descricao': 'Descrição', 'valor_encontrado': 'Valor encontrado', 'valor_esperado': 'Esperado'
        }
        widths = {'aba': 110, 'linha': 60, 'regra': 90, 'severidade': 90, 'descricao': 360, 'valor_encontrado': 170, 'valor_esperado': 170}
        for c in cols:
            self.tree.heading(c, text=headings[c])
            self.tree.column(c, width=widths[c], anchor='w')
        vsb = ttk.Scrollbar(tabela_frame, orient='vertical', command=self.tree.yview)
        hsb = ttk.Scrollbar(tabela_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        tabela_frame.rowconfigure(0, weight=1)
        tabela_frame.columnconfigure(0, weight=1)

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(filetypes=[('Planilhas Excel', '*.xlsx')])
        if caminho:
            self.caminho_arquivo = caminho
            self.lbl_arquivo.config(text=caminho)

    def analisar(self):
        if not self.caminho_arquivo:
            messagebox.showwarning(APP_TITLE, 'Selecione uma planilha .xlsx antes de analisar.')
            return
        try:
            self.resultados = analisar_planilha(self.caminho_arquivo)
            self._renderizar_resultados()
        except Exception as e:
            messagebox.showerror(APP_TITLE, f'Erro ao analisar planilha:\n{e}')

    def _renderizar_resultados(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in self.resultados:
            self.tree.insert('', 'end', values=(r.aba, r.linha, r.regra, r.severidade, r.descricao, r.valor_encontrado, r.valor_esperado))
        total = len(self.resultados)
        por_aba: Dict[str, int] = {}
        for r in self.resultados:
            por_aba[r.aba] = por_aba.get(r.aba, 0) + 1
        resumo = [f'Total de inconsistências: {total}']
        if por_aba:
            resumo.append('Por aba: ' + ' | '.join(f'{k}: {v}' for k, v in sorted(por_aba.items())))
        else:
            resumo.append('Nenhuma inconsistência encontrada.')
        self.lbl_resumo.config(text='\n'.join(resumo))

    def exportar(self):
        if not self.resultados:
            messagebox.showwarning(APP_TITLE, 'Não há resultados para exportar.')
            return
        caminho = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')], initialfile='relatorio_inconsistencias.xlsx')
        if not caminho:
            return
        df = pd.DataFrame([asdict(r) for r in self.resultados])
        df.to_excel(caminho, index=False)
        messagebox.showinfo(APP_TITLE, f'Relatório exportado com sucesso:\n{caminho}')


if __name__ == '__main__':
    root = tk.Tk()
    app = App(root)
    root.mainloop()
