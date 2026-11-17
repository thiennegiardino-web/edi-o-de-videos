# -*- coding: utf-8 -*-
"""Infraestrutura comum do gerador das planilhas (De Malas e Milhas — Milhas do zero)."""
import datetime, re, os
import xlsxwriter

REPO = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
TEXTO = os.path.join(REPO, 'entrega-diagramacao', 'ebook-1-milhas-do-zero', '01-texto-consolidado.md')

# Paleta (marfim claro, cobre/rosé, texto escuro)
IVORY = '#FBF5EC'; INPUT_BG = '#FFFEFA'; CALC_BG = '#F4E3D9'; COPPER = '#A8562B'; COPPER2 = '#C2703D'
ROSE = '#E8B9A5'; ROSE_L = '#F6E4DA'; TEXT = '#3A1F14'; BORDER = '#D9B8A3'; GRAY = '#8A7B73'
RED_L = '#F5C6C0'; ORANGE_L = '#F9D9A8'; YELLOW_L = '#FCEFB4'; YELLOW_LL = '#FEF8D8'; GREEN_L = '#D5E8D0'; GRAY_L = '#E1DDD9'; BLUE_L = '#D3E3F0'

NAMES = {
 1: '01 Comece aqui', 2: '02 Diagnóstico', 3: '03 Contas', 4: '04 Lotes', 5: '05 Acúmulo', 6: '06 Compras',
 7: '07 Comparar', 8: '08 Alterar', 9: '09 Clubes e compra', 10: '10 Transferências', 11: '11 Renovar',
 12: '12 Plano 30 dias', 13: '13 Checklists', 14: '14 Referências'}

def q(n):  # nome da aba entre aspas para fórmulas
    return "'" + NAMES[n] + "'"

def strip_md(s):
    s = re.sub(r'\*\*(.*?)\*\*', r'\1', s)
    s = re.sub(r'(?<!\*)\*(?!\*)(.*?)(?<!\*)\*(?!\*)', r'\1', s)
    return s.strip()

def excel_serial(d):
    if isinstance(d, datetime.datetime):
        d = d.date()
    return (d - datetime.date(1899, 12, 30)).days

class Book:
    def __init__(self, path, examples, values=None):
        self.path = path; self.ex = examples; self.values = values or {}
        self.wb = xlsxwriter.Workbook(path, {'use_future_functions': True, 'strings_to_urls': False})
        self.wb.set_properties({'title': 'Milhas do zero — ' + ('Exemplos preenchidos' if examples else 'Planilha do aluno'),
                                'author': 'De Malas e Milhas', 'company': 'De Malas e Milhas',
                                'comments': 'Material complementar do e-book Milhas do zero. Sem macros nem conexões externas.'})
        self.formula_cells = []   # (aba, célula)
        self._fmt = {}
        self.wb.set_size(1500, 900)

    # ---------- formatos ----------
    def f(self, **kw):
        key = tuple(sorted(kw.items()))
        if key not in self._fmt:
            base = {'font_name': 'Calibri', 'font_size': 10, 'font_color': TEXT, 'valign': 'vcenter'}
            base.update(kw)
            self._fmt[key] = self.wb.add_format(base)
        return self._fmt[key]

    NUM = {'money': '"R$" #,##0.00', 'int': '#,##0', 'pts': '#,##0', 'date': 'dd/mm/yyyy', 'pct': '0.0%', 'dec': '#,##0.00##', 'mes': '0'}

    def inp_fmt(self, kind):
        kw = {'bg_color': INPUT_BG, 'border': 1, 'border_color': BORDER, 'locked': False}
        if kind == 'textnum':   # texto (ex.: número do pedido, com zeros à esquerda)
            kw.update({'num_format': '@', 'align': 'left'})
        elif kind in ('text', 'long'):
            kw.update({'text_wrap': True, 'valign': 'top'} if kind == 'long' else {'text_wrap': True})
        else:
            kw.update({'num_format': self.NUM[kind], 'align': 'right' if kind != 'date' else 'center'})
        return self.f(**kw)

    def calc_fmt(self, kind, bold=False):
        kw = {'bg_color': CALC_BG, 'border': 1, 'border_color': BORDER, 'locked': True, 'italic': False}
        if bold: kw['bold'] = True
        if kind in ('text', 'long'):
            kw.update({'text_wrap': True, 'valign': 'top'} if kind == 'long' else {'text_wrap': True})
        else:
            kw.update({'num_format': self.NUM[kind], 'align': 'right' if kind != 'date' else 'center'})
        return self.f(**kw)

    def sheet(self, n):
        return Sheet(self, n)

class Sheet:
    def __init__(self, book, n):
        self.b = book; self.n = n; self.name = NAMES[n]
        self.ws = book.wb.add_worksheet(self.name)
        self.ws.hide_gridlines(2)
        self.ws.set_tab_color(COPPER2 if n not in (1,) else COPPER)
        self.ws.set_default_row(16)
        self.ws.set_column('A:A', 28)

    # ---------- escrita ----------
    def t(self, cell, text, **fmt):
        self.ws.write(cell, text, self.b.f(**fmt))

    def title(self, titulo, subtitulo, instr=None):
        ws = self.ws
        ws.set_row(0, 26); ws.set_row(1, 18); ws.set_row(2, 16); ws.set_row(3, 16); ws.set_row(4, 16)
        self.t('A1', titulo, font_name='Georgia', font_size=18, bold=True, font_color=COPPER)
        self.t('A2', subtitulo, font_name='Georgia', font_size=11, italic=True, font_color=TEXT)
        if instr:
            self.t('A3', instr, font_size=10, font_color=TEXT)
        # link de volta
        return

    def back_link(self, cell):
        self.ws.write_url(cell, "internal:'%s'!A1" % NAMES[1], self.b.f(font_color=COPPER, underline=1, font_size=10),
                          string='← Voltar à aba 01 Comece aqui', tip='Ir para a aba 01 Comece aqui')

    def head(self, row, headers, col0=0, height=45, calc_cols=()):
        ws = self.ws
        ws.set_row(row - 1, height)
        for i, h in enumerate(headers):
            colr = COPPER2 if (col0 + i) in calc_cols else COPPER
            ws.write(row - 1, col0 + i, h, self.b.f(bold=True, font_color='#FFFFFF', bg_color=colr, text_wrap=True, border=1,
                                                       border_color='#FFFFFF', align='center', valign='vcenter'))

    def section(self, cell, text, ncols=1):
        self.ws.write(cell, text, self.b.f(bold=True, font_color=COPPER, font_size=11, bg_color=ROSE_L, bottom=2, bottom_color=COPPER2))

    def label(self, cell, text, bold=False, wrap=True, **kw):
        self.ws.write(cell, text, self.b.f(bold=bold, text_wrap=wrap, valign='top' if wrap else 'vcenter', **kw))

    def inp(self, cell, kind='text', ex=None):
        fmt = self.b.inp_fmt(kind)
        if self.b.ex and ex is not None:
            if isinstance(ex, (datetime.date, datetime.datetime)):
                self.ws.write_datetime(cell, datetime.datetime(ex.year, ex.month, ex.day), fmt)
            elif kind == 'textnum':
                self.ws.write_string(cell, str(ex), fmt)
            else:
                self.ws.write(cell, ex, fmt)
        else:
            self.ws.write_blank(cell, None, fmt)

    def calc(self, cell, formula, kind='text', bold=False):
        fmt = self.b.calc_fmt(kind, bold)
        key = (self.name, cell)
        self.b.formula_cells.append(key)
        v = self.b.values.get(key, None)
        if v is None:
            self.ws.write_formula(cell, formula, fmt)
        else:
            self.ws.write_formula(cell, formula, fmt, v)

    def static_calc(self, cell, text, kind='text'):
        self.ws.write(cell, text, self.b.calc_fmt(kind))

    def validate_list(self, rng, items=None, source=None, msg=None, strict=True):
        opt = {'validate': 'list', 'source': items if items is not None else source, 'ignore_blank': True,
               'error_type': 'stop' if strict else 'warning',
               'error_title': 'Valor não previsto', 'error_message': 'Escolha uma das opções da lista.' if strict else 'Valor diferente da lista. Se for outro, descreva também em Observações.'}
        if msg: opt.update({'input_title': 'Dica', 'input_message': msg})
        self.ws.data_validation(rng, opt)

    def validate_num(self, rng, minimum=0, maximum=None, integer=False, msg=None):
        opt = {'validate': 'integer' if integer else 'decimal', 'criteria': 'between' if maximum is not None else '>=',
               'minimum': minimum, 'ignore_blank': True, 'error_title': 'Número inválido',
               'error_message': 'Digite um número ' + ('inteiro ' if integer else '') + ('maior ou igual a %s.' % minimum if maximum is None else 'entre %s e %s.' % (minimum, maximum))}
        if maximum is not None: opt['maximum'] = maximum
        if msg: opt.update({'input_title': 'Dica', 'input_message': msg})
        self.ws.data_validation(rng, opt)

    def validate_date(self, rng, msg=None):
        opt = {'validate': 'date', 'criteria': 'between', 'minimum': datetime.date(2000, 1, 1), 'maximum': datetime.date(2100, 12, 31),
               'ignore_blank': True, 'error_title': 'Data inválida', 'error_message': 'Digite uma data válida (dd/mm/aaaa).'}
        if msg: opt.update({'input_title': 'Dica', 'input_message': msg})
        self.ws.data_validation(rng, opt)

    def cf(self, rng, formula, fmt_kw):
        opts = {'type': 'formula', 'criteria': formula, 'format': self.b.wb.add_format(fmt_kw)}
        parts = rng.split(',')
        if len(parts) > 1:
            opts['multi_range'] = ' '.join(parts)
        self.ws.conditional_format(parts[0], opts)

    def finish(self, header_row=None, area=None, landscape=True, freeze=None, autofilter=None, protect=True, fit_w=1):
        ws = self.ws
        ws.set_paper(9)
        if landscape: ws.set_landscape()
        ws.fit_to_pages(fit_w, 0)
        ws.set_margins(left=0.4, right=0.4, top=0.6, bottom=0.6)
        ws.set_header('&L&"Calibri,Bold"De Malas e Milhas&R&"Calibri"Milhas do zero — ' + self.name)
        ws.set_footer('&L&"Calibri"Material complementar do e-book. Confira regras, prazos e validades no site oficial.&R&"Calibri"Página &P de &N')
        if header_row: ws.repeat_rows(header_row - 1)
        if area: ws.print_area(area)
        if freeze: ws.freeze_panes(*freeze)
        if autofilter: ws.autofilter(autofilter)
        if protect:
            ws.protect('', {'autofilter': True, 'sort': False, 'format_columns': True, 'format_rows': True,
                            'select_locked_cells': True, 'select_unlocked_cells': True})

def parse_checklists():
    """Lê todos os itens '- [ ]' do texto aprovado, com capítulo, subtítulo e grupo."""
    lines = open(TEXTO, encoding='utf-8').read().split('\n')
    cap = None; h3 = None; bold = None; out = []
    for l in lines:
        m = re.match(r'^## Capítulo (\d+) — (.*)', l)
        if m: cap = int(m.group(1)); h3 = None; bold = None; continue
        m = re.match(r'^### (.*)', l)
        if m: h3 = strip_md(m.group(1)); bold = None; continue
        m = re.match(r'^\*\*([^*]+)\*\*\s*$', l)
        if m: bold = m.group(1).strip(); continue
        m = re.match(r'^- \[ \] (.*)', l)
        if m: out.append({'cap': cap, 'h3': h3, 'bold': bold, 'raw': m.group(1), 'text': strip_md(m.group(1))})
    return out

def parse_table(start_pattern, occurrence=1):
    """Devolve as linhas (listas de células) da tabela cuja linha de cabeçalho contém start_pattern."""
    lines = open(TEXTO, encoding='utf-8').read().split('\n')
    seen = 0
    for i, l in enumerate(lines):
        t = l.lstrip('> ').rstrip()
        if t.startswith('|') and start_pattern in t:
            seen += 1
            if seen == occurrence:
                rows = []
                j = i
                while j < len(lines) and lines[j].lstrip('> ').startswith('|'):
                    cells = [strip_md(c.strip()) for c in lines[j].lstrip('> ').strip().strip('|').split('|')]
                    rows.append(cells); j += 1
                return [r for k, r in enumerate(rows) if k != 1]  # remove a linha de separação
    raise KeyError(start_pattern)

# ---------- Layout (linhas) compartilhado entre as abas ----------
LOTES = (7, 156)      # 150 linhas de lotes
CONTAS = (7, 36)      # 30 contas
COMPRAS = (7, 106)    # 100 compras
ACUMULO = (7, 36)     # 30 linhas
DIAG_PROG = (36, 50)  # programas do diagnóstico (ajustado no código da aba)
REF_DATE = datetime.date(2026, 11, 16)

def rng(n, col, r0, r1):
    return "%s!$%s$%d:$%s$%d" % (q(n), col, r0, col, r1)


def parse_cap12_questions():
    """As sete perguntas do Capítulo 12 ("**1. Quanto valor eu obtenho?**" ...)."""
    lines = open(TEXTO, encoding='utf-8').read().split('\n'); cap = None; out = []
    for l in lines:
        m = re.match(r'^## Capítulo (\d+)', l)
        if m: cap = int(m.group(1)); continue
        if cap == 12:
            m = re.match(r'^\*\*(\d)\. (.*)\*\*\s*$', l)
            if m: out.append(strip_md(m.group(2)))
    return out
