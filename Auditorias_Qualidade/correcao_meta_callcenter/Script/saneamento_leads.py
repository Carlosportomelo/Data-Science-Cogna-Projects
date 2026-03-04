"""Saneamento / Validação de Cobertura (HubSpot vs Meta)

Requisitos implementados:
- Localiza automaticamente o arquivo do HubSpot contendo "Novos Negócios" em Data/
- Lê XLS/XLSX e CSV com tentativas de encodings (utf-8, latin1, utf-16) e separadores comuns
- Normaliza e remove duplicatas pela chave de e-mail
- Lê todos os arquivos que começam com "nbfo" em Data/Pastas meta/
- Tenta ler primeiro como UTF-16 com tabulação como pedido, com fallbacks
- Detecta dinamicamente colunas de e-mail e data de nascimento
- Realiza LEFT JOIN mantendo HubSpot à esquerda e marca ENCONTRADO / NÃO ENCONTRADO
- Salva Relatorio_Cobertura.csv na raiz com separador ';'
"""

from typing import List, Optional
import pandas as pd
import numpy as np
import re
import glob
import os
import sys

# ================= CONFIGURAÇÕES DE DIRETÓRIOS =================
BASE_DIR = os.getcwd()
DATA_DIR = os.path.join(BASE_DIR, 'Data')
META_DIR = os.path.join(DATA_DIR, 'Pastas meta')
OUTPUT_DIR = os.path.join(BASE_DIR, 'Output')

os.makedirs(OUTPUT_DIR, exist_ok=True)

# ================= UTILITÁRIOS =================
EMAIL_COL_KEYWORDS = ['email', 'e-mail', 'mail']
BIRTH_COL_KEYWORDS = ['birth', 'nasc', 'idade', 'data_nasc', 'nascimento', 'birth_date']


def clean_email(email: object) -> str:
    """Normaliza e-mail: string baixa, sem espaços"""
    if pd.isna(email):
        return ""
    return str(email).lower().strip()


def try_read_csv_with_fallback(path: str, verbose: bool = True) -> Optional[pd.DataFrame]:
    """Tenta ler um CSV/TSV com combinações de encodings e separadores comuns.
    Prioriza UTF-16 + tab (Meta) como pedido.
    """
    encodings = ['utf-16', 'utf-8', 'latin1']
    separators = ['\t', ',', ';']

    last_exc = None
    for enc in encodings:
        for sep in separators:
            try:
                df = pd.read_csv(path, sep=sep, encoding=enc)
                return df
            except Exception as e:
                last_exc = e
                continue
    # Último recurso: deixar pandas tentar autodetect
    try:
        return pd.read_csv(path)
    except Exception as e:
        # Falhou completamente
        if verbose:
            print(f"⚠️ Falha ao ler {os.path.basename(path)}: {e}")
        return None


def find_hubspot_file(data_dir: str, preferred_name: Optional[str] = None) -> Optional[str]:
    """Encontra arquivo do HubSpot. Se preferred_name for passado, tenta casar exatamente ou por substring.
    Caso contrário, busca por padrões "Novos Negócios" e, em último caso, qualquer CSV/XLS.
    """
    # Se o usuário especificou um arquivo exato ou parcial, tenta achar primeiro
    if preferred_name:
        # procura por correspondência exata (basename)
        candidate_exact = os.path.join(data_dir, preferred_name)
        if os.path.isfile(candidate_exact):
            return candidate_exact
        # procura por nomes que contenham a string fornecida
        candidates = glob.glob(os.path.join(data_dir, '**', f'*{preferred_name}*'), recursive=True)
        if candidates:
            return max(candidates, key=os.path.getctime)

    # Busca por nomes que contenham 'Novos Negócios' (case-insensitive)
    candidates = glob.glob(os.path.join(data_dir, '**', '*Novos Negócios*'), recursive=True)
    if not candidates:
        # fallback para qualquer CSV/XLSX começando com 'novos'
        candidates = glob.glob(os.path.join(data_dir, '**', '*'), recursive=True)
        candidates = [c for c in candidates if os.path.isfile(c) and os.path.basename(c).lower().startswith(('novos',))]
    if not candidates:
        candidates = glob.glob(os.path.join(data_dir, '**', '*Novos*'), recursive=True)

    if not candidates:
        # try any csv/xlsx
        candidates = glob.glob(os.path.join(data_dir, '**', '*.csv'), recursive=True) + \
                     glob.glob(os.path.join(data_dir, '**', '*.xls'), recursive=True) + \
                     glob.glob(os.path.join(data_dir, '**', '*.xlsx'), recursive=True)

    if not candidates:
        return None
    # retorna o mais recente
    return max(candidates, key=os.path.getctime)


def read_hubspot(file_path: str) -> pd.DataFrame:
    print(f"📖 Lendo HubSpot: {os.path.basename(file_path)}")
    df = None
    _, ext = os.path.splitext(file_path.lower())

    if ext in ('.xls', '.xlsx'):
        # Pandas read_excel é robusto
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            # Tentar ler como HTML (algumas exportações do HubSpot são HTML salvo em .xls)
            try:
                df = pd.read_html(file_path)[0]
            except Exception as e2:
                raise RuntimeError(f"Erro ao ler arquivo Excel do HubSpot: {e} / {e2}")
    else:
        # Tenta diferentes encodings e separadores
        df = try_read_csv_with_fallback(file_path)
        if df is None:
            raise RuntimeError("Não foi possível ler o CSV do HubSpot com os encodings/separadores testados.")

    # Identificar coluna de e-mail
    cols = list(df.columns)
    col_email = next((c for c in cols if any(k in str(c).lower() for k in EMAIL_COL_KEYWORDS)), None)
    if col_email is None:
        raise RuntimeError(f"Coluna de e-mail não encontrada. Colunas: {cols}")

    df['key_email'] = df[col_email].apply(clean_email)
    df = df[df['key_email'] != ""].drop_duplicates(subset=['key_email']).reset_index(drop=True)
    return df


def read_meta_all(meta_dir: str) -> pd.DataFrame:
    print(f"📥 Procurando arquivos 'nbfo*' em: {meta_dir} (recursivo)")
    if not os.path.isdir(meta_dir):
        raise RuntimeError(f"Pasta Meta não encontrada: {meta_dir}")

    # Captura todos os arquivos que comecem com nbfo (qualquer extensão)
    files = glob.glob(os.path.join(meta_dir, '**', 'nbfo*'), recursive=True)
    files = [f for f in files if os.path.isfile(f)]
    
    if not files:
        # Fallback para busca genérica e filtro manual
        all_files = glob.glob(os.path.join(meta_dir, '**', '*'), recursive=True)
        files = [f for f in all_files if os.path.isfile(f) and os.path.basename(f).lower().startswith('nbfo')]

    if not files:
        raise RuntimeError("Nenhum arquivo 'nbfo*' encontrado na pasta Meta.")

    dfs: List[pd.DataFrame] = []
    for f in files:
        df_raw = try_read_csv_with_fallback(f)
        if df_raw is None:
            continue
        # Normaliza colunas
        df_raw.columns = [str(c).lower().strip() for c in df_raw.columns]
        cols = df_raw.columns.tolist()
        col_email = next((c for c in cols if 'email' in c), None)
        col_birth = next((c for c in cols if any(k in c for k in BIRTH_COL_KEYWORDS)), None)
        if col_email is None or col_birth is None:
            # Pula arquivos que não tenham as colunas necessárias
            print(f"⚠️ Arquivo {os.path.basename(f)} não tem colunas necessárias (email/data nascimento). Pulando.")
            continue

        temp = df_raw[[col_email, col_birth]].copy()
        temp.columns = ['email_raw', 'data_nascimento_meta']
        temp['key_email'] = temp['email_raw'].apply(clean_email)
        # Filtra chaves vazias e duplicatas
        temp = temp[temp['key_email'] != ""].drop_duplicates(subset=['key_email'])
        dfs.append(temp[['key_email', 'data_nascimento_meta']])

    if not dfs:
        raise RuntimeError("Nenhum registro válido extraído dos arquivos Meta.")

    df_concat = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=['key_email']).reset_index(drop=True)
    return df_concat


# ----------------- Novas utilidades para matching avançado -----------------
PHONE_COL_KEYWORDS = ['phone', 'telefone', 'cel', 'celular', 'mobile', 'telefone1', 'telefone2']

def clean_phone(phone: object) -> str:
    """Remove tudo que não é dígito e retorna string curta (apenas números)."""
    if pd.isna(phone):
        return ""
    s = re.sub(r"\D", "", str(phone))
    return s  # mantemos o número inteiro sem remover zeros aqui


def normalize_phone(phone_digits: str) -> str:
    """Normaliza telefone para uma forma canônica útil para comparação.
    Estratégia:
    - remove não-dígitos (suportado fora)
    - retorna a versão com os últimos 11 dígitos se disponível, senão últimos 10, senão original
    """
    if not phone_digits:
        return ""
    s = re.sub(r"\D", "", str(phone_digits))
    s = s.lstrip('+').lstrip('00')
    # remove leading zeros só para normalizar formatos locais
    s = s.lstrip('0')
    # prefere 11 (DDD+9) se disponível, senão 10
    if len(s) >= 11:
        return s[-11:]
    if len(s) >= 10:
        return s[-10:]
    return s


def phones_match(a: str, b: str, min_suffix_len: int = 8) -> bool:
    """Retorna True se os telefones combinarem por normalização ou sufixo (mínimo min_suffix_len dígitos)."""
    if not a or not b:
        return False
    a_norm = normalize_phone(a)
    b_norm = normalize_phone(b)
    if not a_norm or not b_norm:
        return False
    if a_norm == b_norm:
        return True
    # sufix match: checa se um termina com o outro e que o menor tenha ao menos min_suffix_len
    if len(a_norm) >= min_suffix_len and a_norm.endswith(b_norm[-min_suffix_len:]):
        return True
    if len(b_norm) >= min_suffix_len and b_norm.endswith(a_norm[-min_suffix_len:]):
        return True
    # última tentativa: compara últimos 8 dígitos
    return a_norm.endswith(b_norm[-8:]) or b_norm.endswith(a_norm[-8:])


from difflib import SequenceMatcher

def fuzzy_score(a: str, b: str) -> float:
    return SequenceMatcher(None, a or '', b or '').ratio()


# ----------------- NOVO: Tratamento de idade ou data de nascimento -----------------

def compute_age_from_dob(dob: pd.Timestamp) -> Optional[int]:
    if pd.isna(dob):
        return None
    today = pd.Timestamp.now().normalize()
    years = today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    return int(years)


def parse_age_or_dob(value: object) -> (Optional[pd.Timestamp], Optional[int]):
    """Tenta interpretar um valor como data de nascimento ou idade.
    Retorna (dob_datetime_or_None, age_int_or_None).
    Regras:
    - Se parse como data válido -> retorna (datetime, computed_age)
    - Se numérico e entre 5-120 -> retorna (None, int)
    - Se ano (4 dígitos plausível) -> retorna (datetime Jan1 year, computed_age)
    - Else -> (None, None)
    """
    if pd.isna(value) or (isinstance(value, str) and value.strip() == ''):
        return (None, None)
    s = str(value).strip()

    # Tenta parse de data
    try:
        dt = pd.to_datetime(s, dayfirst=True, errors='coerce')
        if not pd.isna(dt):
            age = compute_age_from_dob(dt)
            return (dt, age)
    except Exception:
        pass

    # Extrai primeiro número encontrado
    m = re.search(r"(\d{4})", s)
    if m:
        year = int(m.group(1))
        if 1900 <= year <= pd.Timestamp.now().year:
            dt = pd.to_datetime(f"{year}-01-01")
            age = compute_age_from_dob(dt)
            return (dt, age)

    m2 = re.search(r"(\d{1,3})", s)
    if m2:
        num = int(m2.group(1))
        if 5 <= num <= 120:
            return (None, int(num))

    return (None, None)


# Atualiza funções de leitura para capturar telefone quando possível
def read_hubspot(file_path: str) -> pd.DataFrame:
    print(f"📖 Lendo HubSpot: {os.path.basename(file_path)}")
    df = None
    _, ext = os.path.splitext(file_path.lower())

    if ext in ('.xls', '.xlsx'):
        # Pandas read_excel é robusto
        try:
            df = pd.read_excel(file_path)
        except Exception as e:
            # Tentar ler como HTML (algumas exportações do HubSpot são HTML salvo em .xls)
            try:
                df = pd.read_html(file_path)[0]
            except Exception as e2:
                raise RuntimeError(f"Erro ao ler arquivo Excel do HubSpot: {e} / {e2}")
    else:
        # Tenta diferentes encodings e separadores
        df = try_read_csv_with_fallback(file_path)
        if df is None:
            raise RuntimeError("Não foi possível ler o CSV do HubSpot com os encodings/separadores testados.")

    # Identificar coluna de e-mail
    cols = list(df.columns)
    col_email = next((c for c in cols if any(k in str(c).lower() for k in EMAIL_COL_KEYWORDS)), None)
    if col_email is None:
        raise RuntimeError(f"Coluna de e-mail não encontrada. Colunas: {cols}")

    # Detecta coluna de telefone (opcional)
    col_phone = next((c for c in cols if any(k in str(c).lower() for k in PHONE_COL_KEYWORDS)), None)

    df['key_email'] = df[col_email].apply(clean_email)
    if col_phone:
        df['phone_hs'] = df[col_phone].apply(clean_phone)
        df['phone_hs_norm'] = df[col_phone].apply(lambda x: normalize_phone(clean_phone(x)))
    else:
        df['phone_hs'] = ""
        df['phone_hs_norm'] = ""

    df = df[df['key_email'] != ""].drop_duplicates(subset=['key_email']).reset_index(drop=True)
    return df


def read_meta_all(meta_dir: str) -> pd.DataFrame:
    print(f"📥 Procurando arquivos 'nbfo*' em: {meta_dir} (recursivo)")
    if not os.path.isdir(meta_dir):
        raise RuntimeError(f"Pasta Meta não encontrada: {meta_dir}")

    # Captura todos os arquivos que comecem com nbfo (qualquer extensão)
    files = glob.glob(os.path.join(meta_dir, '**', 'nbfo*'), recursive=True)
    files = [f for f in files if os.path.isfile(f)]
    
    if not files:
        # Fallback para busca genérica e filtro manual
        all_files = glob.glob(os.path.join(meta_dir, '**', '*'), recursive=True)
        files = [f for f in all_files if os.path.isfile(f) and os.path.basename(f).lower().startswith('nbfo')]

    if not files:
        raise RuntimeError("Nenhum arquivo 'nbfo*' encontrado na pasta Meta.")

    dfs: List[pd.DataFrame] = []
    for f in files:
        df_raw = try_read_csv_with_fallback(f)
        if df_raw is None:
            continue
        # Normaliza colunas
        df_raw.columns = [str(c).lower().strip() for c in df_raw.columns]
        cols = df_raw.columns.tolist()
        col_email = next((c for c in cols if 'email' in c), None)
        col_birth = next((c for c in cols if any(k in c for k in BIRTH_COL_KEYWORDS)), None)
        col_phone = next((c for c in cols if any(k in c for k in PHONE_COL_KEYWORDS)), None)
        if col_email is None or col_birth is None:
            # Pula arquivos que não tenham as colunas necessárias
            print(f"⚠️ Arquivo {os.path.basename(f)} não tem colunas necessárias (email/data nascimento). Pulando.")
            continue

        temp = df_raw[[col_email, col_birth]].copy()
        temp.columns = ['email_raw', 'data_nascimento_meta']
        if col_phone:
            temp['phone_meta'] = df_raw[col_phone].apply(clean_phone)
            temp['phone_meta_norm'] = df_raw[col_phone].apply(lambda x: normalize_phone(clean_phone(x)))
        else:
            temp['phone_meta'] = ""
            temp['phone_meta_norm'] = ""

        temp['key_email'] = temp['email_raw'].apply(clean_email)
        # Parseia campo que pode ser data ou idade
        parsed = temp['data_nascimento_meta'].apply(lambda x: parse_age_or_dob(x))
        temp['data_nascimento_meta'] = parsed.apply(lambda x: x[0])
        temp['idade_meta'] = parsed.apply(lambda x: x[1])
        # adicione metadados de origem (arquivo / campanha)
        temp['source_file'] = os.path.basename(f)
        # tentativa simples de extrair campanha do nome do arquivo (parte até primeiro underscore ou o nome todo)
        temp['campaign'] = temp['source_file'].apply(lambda nm: nm.split('_')[0] if '_' in nm else os.path.splitext(nm)[0])
        # Filtra chaves vazias e duplicatas
        temp = temp[temp['key_email'] != ""].drop_duplicates(subset=['key_email'])
        dfs.append(temp[['key_email', 'data_nascimento_meta', 'idade_meta', 'phone_meta', 'phone_meta_norm', 'source_file', 'campaign']])

    if not dfs:
        raise RuntimeError("Nenhum registro válido extraído dos arquivos Meta.")

    df_concat = pd.concat(dfs, ignore_index=True).drop_duplicates(subset=['key_email']).reset_index(drop=True)
    return df_concat


def find_possible_matches(df_hs: pd.DataFrame, df_meta: pd.DataFrame, top_n: int = 3, min_phone_len: int = 8) -> pd.DataFrame:
    """Para cada lead do HubSpot que não foi encontrado por e-mail exato, tenta:
    - match por telefone (normalizado/sufixo)
    - fuzzy matching de e-mail (com score)
    Retorna um DataFrame com possíveis matches ordenados por score, já limpo e deduplicado.
    """
    not_found = df_hs[~df_hs['key_email'].isin(df_meta['key_email'])].copy()
    meta_emails = df_meta['key_email'].tolist()

    results = []
    for _, row in not_found.iterrows():
        hub_email = row['key_email']
        hub_phone = row.get('phone_hs', '')
        hub_phone_norm = normalize_phone(hub_phone) if hub_phone else ''

        # 1) Phone match (com normalização/sufixo)
        if hub_phone_norm:
            mask = df_meta['phone_meta_norm'].apply(lambda x: phones_match(hub_phone_norm, x))
            matched = df_meta[mask]
            for _, m in matched.iterrows():
                results.append({
                    'hub_email': hub_email,
                    'hub_phone': hub_phone,
                    'hub_phone_norm': hub_phone_norm,
                    'meta_email': m['key_email'],
                    'meta_phone': m.get('phone_meta', ''),
                    'meta_phone_norm': m.get('phone_meta_norm', ''),
                    'meta_data_nascimento': m.get('data_nascimento_meta', ''),
                    'method': 'phone_exact',
                    'score': 1.0
                })
            if not matched.empty:
                continue  # já temos match por telefone, pula outras tentativas

        # 2) Fuzzy on email (limitado a top_n melhores)
        scores = []
        for meta_email in meta_emails:
            sc = fuzzy_score(hub_email, meta_email)
            if sc > 0.6:  # filtro inicial para performance
                scores.append((meta_email, sc))
        scores.sort(key=lambda x: x[1], reverse=True)
        for meta_email, sc in scores[:top_n]:
            mrow = df_meta[df_meta['key_email'] == meta_email].iloc[0]
            results.append({
                'hub_email': hub_email,
                'hub_phone': hub_phone,
                'hub_phone_norm': hub_phone_norm,
                'meta_email': meta_email,
                'meta_phone': mrow.get('phone_meta', ''),
                'meta_phone_norm': mrow.get('phone_meta_norm', ''),
                'meta_data_nascimento': mrow.get('data_nascimento_meta', ''),
                'method': 'email_fuzzy',
                'score': sc
            })

        # 3) Domain-aware local-part fuzzy (se já não teve matches fortes)
        # (ex: same domain and high local-part similarity)
        if not any(r['hub_email'] == hub_email and r['score'] >= 0.85 for r in results):
            hub_local, _, hub_dom = hub_email.partition('@')
            domain_candidates = [e for e in meta_emails if e.endswith('@' + hub_dom)] if hub_dom else []
            for meta_email in domain_candidates:
                local_meta = meta_email.partition('@')[0]
                sc_local = fuzzy_score(hub_local, local_meta)
                if sc_local > 0.7:
                    mrow = df_meta[df_meta['key_email'] == meta_email].iloc[0]
                    results.append({
                        'hub_email': hub_email,
                        'hub_phone': hub_phone,
                        'hub_phone_norm': hub_phone_norm,
                        'meta_email': meta_email,
                        'meta_phone': mrow.get('phone_meta', ''),
                        'meta_phone_norm': mrow.get('phone_meta_norm', ''),
                        'meta_data_nascimento': mrow.get('data_nascimento_meta', ''),
                        'method': 'domain_local_fuzzy',
                        'score': sc_local
                    })

    df_poss = pd.DataFrame(results)
    if df_poss.empty:
        return df_poss

    # Filtra telefones inválidos/curtos (ex: valores corrompidos como '3')
    def valid_phone_row(row):
        mpn = str(row.get('meta_phone_norm', '') or '')
        # se existir telefone meta, requer tamanho mínimo
        if mpn and len(mpn) < min_phone_len:
            return False
        return True

    df_poss = df_poss[df_poss.apply(valid_phone_row, axis=1)].copy()

    # Remove duplicações triviais e ordena por score
    df_poss = df_poss.drop_duplicates(subset=['hub_email', 'meta_email', 'method'])
    df_poss = df_poss.sort_values(by=['hub_email', 'score'], ascending=[True, False]).reset_index(drop=True)

    # Adiciona ranking por hub_email (1 = melhor candidato)
    df_poss['rank'] = df_poss.groupby('hub_email')['score'].rank(method='first', ascending=False).astype(int)

    # remove repetições por meta_phone inválido (ex: '3') e normalize string fields
    df_poss = df_poss[df_poss['meta_phone'].apply(lambda x: str(x).isdigit() and len(str(x))>=8 if x else True)].copy()
    for c in ['hub_email', 'meta_email', 'hub_phone', 'meta_phone']:
        if c in df_poss.columns:
            df_poss[c] = df_poss[c].astype(str).str.strip()

    return df_poss


# ----------------- FUNÇÃO: Preenchimento exato para HubSpot -----------------
def apply_exact_fill(df_final: pd.DataFrame, df_meta: pd.DataFrame) -> (pd.DataFrame, pd.DataFrame):
    """Tenta preencher 'data_nascimento_meta' no DataFrame resultante usando:
    - correspondência exata por e-mail (já feita no merge)
    - correspondência exata por telefone (preenche onde e-mail não achou)
    Retorna (df_final_com_preenchimentos, df_relatorio_de_mudancas)
    """
    df = df_final.copy()
    # Garante colunas de tracking
    df['filled_by'] = df.get('filled_by', '')
    df['meta_matched_email'] = df.get('meta_matched_email', '')
    df['meta_matched_phone'] = df.get('meta_matched_phone', '')

    changes = []
    for idx, row in df.iterrows():
        # se já tem data (email exato), marca filled_by=email e computa idade se possível
        if pd.notna(row.get('data_nascimento_meta')):
            if row['Status'] == 'ENCONTRADO':
                df.at[idx, 'filled_by'] = 'email'
                df.at[idx, 'meta_matched_email'] = row.get('key_email', '')
                # calcula idade a partir de data
                try:
                    age = compute_age_from_dob(row.get('data_nascimento_meta'))
                    if age is not None:
                        df.at[idx, 'idade'] = age
                except Exception:
                    pass
            continue

        # Se já tem idade meta (vinda da base Meta), preenche direto
        hub_phone = row.get('phone_hs', '')
        # prioridade: se meta tem idade pelo email exato (merge), usa
        meta_age = row.get('idade_meta') if 'idade_meta' in row else None
        if pd.notna(meta_age) and meta_age != '':
            df.at[idx, 'idade'] = int(meta_age)
            df.at[idx, 'Status'] = 'ENCONTRADO'
            df.at[idx, 'filled_by'] = 'idade_meta'
            df.at[idx, 'meta_matched_email'] = row.get('key_email', '')
            changes.append({
                'hub_email': row.get('key_email', ''),
                'hub_phone': hub_phone,
                'old_field': 'idade',
                'old_value': '',
                'new_field': 'idade',
                'new_value': int(meta_age),
                'method': 'email_exact_idade',
                'meta_email': row.get('key_email', ''),
                'meta_phone': row.get('phone_meta', '') if 'phone_meta' in row else ''
            })
            continue

        if hub_phone:
            hub_phone_norm = normalize_phone(hub_phone)
            mask = df_meta['phone_meta_norm'].apply(lambda x: phones_match(hub_phone_norm, x))
            matched = df_meta[mask]
            if not matched.empty:
                m = matched.iloc[0]
                # Prefer idade_meta if disponível, senão data de nascimento
                meta_age = m.get('idade_meta', None)
                meta_dob = m.get('data_nascimento_meta', None)
                new_age = None
                if pd.notna(meta_age) and meta_age != '':
                    new_age = int(meta_age)
                elif pd.notna(meta_dob):
                    new_age = compute_age_from_dob(meta_dob)
                if new_age is None:
                    continue
                old = row.get('data_nascimento_meta', '')
                df.at[idx, 'idade'] = new_age
                df.at[idx, 'Status'] = 'ENCONTRADO (phone)'
                df.at[idx, 'filled_by'] = 'phone'
                df.at[idx, 'meta_matched_email'] = m.get('key_email', '')
                df.at[idx, 'meta_matched_phone'] = m.get('phone_meta', '')
                changes.append({
                    'hub_email': row.get('key_email', ''),
                    'hub_phone': hub_phone,
                    'old_field': 'idade',
                    'old_value': old,
                    'new_field': 'idade',
                    'new_value': new_age,
                    'method': 'phone_exact',
                    'meta_email': m.get('key_email', ''),
                    'meta_phone': m.get('phone_meta', '')
                })
    df_changes = pd.DataFrame(changes)
    return df, df_changes


# ----------------- FUNÇÃO: Resgatar Não Encontrados -----------------
def rescue_nao_encontrados(nao_file: str, df_meta: pd.DataFrame, fuzzy_threshold: float = 0.85) -> pd.DataFrame:
    """Lê um CSV com colunas 'E-mail' e 'phone_hs' (padrão `Relatorio_NaoEncontrados.csv`) e tenta resgatar idade/data
    Retorna DataFrame com colunas: hub_email, hub_phone, method, matched, meta_email, meta_phone, meta_dob, meta_idade, inferred_age, score
    """
    df_nf = pd.read_csv(nao_file, sep=';', encoding='utf-8-sig')
    results = []
    meta_emails = df_meta['key_email'].tolist()

    for _, r in df_nf.iterrows():
        hub_email = str(r['E-mail']).strip()
        hub_phone = str(r.get('phone_hs', '') or '').strip()
        matched = False
        out = {
            'hub_email': hub_email,
            'hub_phone': hub_phone,
            'method': None,
            'matched': False,
            'meta_email': None,
            'meta_phone': None,
            'meta_dob': None,
            'meta_idade': None,
            'inferred_age': None,
            'score': None
        }

        # 1) email exact
        if hub_email in df_meta['key_email'].values:
            m = df_meta[df_meta['key_email'] == hub_email].iloc[0]
            out.update({
                'method': 'email_exact',
                'matched': True,
                'meta_email': m.get('key_email', None),
                'meta_phone': m.get('phone_meta', None),
                'meta_dob': m.get('data_nascimento_meta', None),
                'meta_idade': int(m.get('idade_meta')) if 'idade_meta' in m.index and not pd.isna(m.get('idade_meta')) and m.get('idade_meta') != '' else None,
                'inferred_age': None,
                'score': 1.0
            })
            # infer age
            if out['meta_idade'] is None and pd.notna(out['meta_dob']):
                out['inferred_age'] = compute_age_from_dob(out['meta_dob'])
            else:
                out['inferred_age'] = out['meta_idade']
            results.append(out)
            continue

        # 2) phone match
        if hub_phone:
            hub_phone_norm = normalize_phone(hub_phone)
            mask = df_meta['phone_meta_norm'].apply(lambda x: phones_match(hub_phone_norm, x))
            matched_rows = df_meta[mask]
            if not matched_rows.empty:
                m = matched_rows.iloc[0]
                out.update({
                    'method': 'phone_match',
                    'matched': True,
                    'meta_email': m.get('key_email', None),
                    'meta_phone': m.get('phone_meta', None),
                    'meta_dob': m.get('data_nascimento_meta', None),
                    'meta_idade': int(m.get('idade_meta')) if 'idade_meta' in m.index and not pd.isna(m.get('idade_meta')) and m.get('idade_meta') != '' else None,
                    'score': 1.0,
                    'source_file': m.get('source_file', None),
                    'campaign': m.get('campaign', None)
                })
                out['inferred_age'] = out['meta_idade'] if out['meta_idade'] is not None else (compute_age_from_dob(out['meta_dob']) if pd.notna(out['meta_dob']) else None)
                results.append(out)
                continue

        # 3) fuzzy email
        best = (None, 0.0)
        for me in meta_emails:
            sc = fuzzy_score(hub_email, me)
            if sc > best[1]:
                best = (me, sc)
        if best[1] >= fuzzy_threshold:
            m = df_meta[df_meta['key_email'] == best[0]].iloc[0]
            out.update({
                'method': 'email_fuzzy',
                'matched': True,
                'meta_email': m.get('key_email', None),
                'meta_phone': m.get('phone_meta', None),
                'meta_dob': m.get('data_nascimento_meta', None),
                'meta_idade': int(m.get('idade_meta')) if 'idade_meta' in m.index and not pd.isna(m.get('idade_meta')) and m.get('idade_meta') != '' else None,
                'score': float(best[1]),
                'source_file': m.get('source_file', None),
                'campaign': m.get('campaign', None)
            })
            out['inferred_age'] = out['meta_idade'] if out['meta_idade'] is not None else (compute_age_from_dob(out['meta_dob']) if pd.notna(out['meta_dob']) else None)
            results.append(out)
            continue

        # não achou
        results.append(out)

    df_res = pd.DataFrame(results)
    # normalize types
    df_res['matched'] = df_res['matched'].astype(bool)
    return df_res


def main():
    import argparse

    parser = argparse.ArgumentParser(description='Validação de cobertura HubSpot vs Meta (matching avançado)')
    parser.add_argument('--hub-file', dest='hub_file', help="Nome exato ou substring do arquivo HubSpot em Data/ a ser usado (ex: 'Novos Negócios 04_02 08h53 2')", default=None)
    parser.add_argument('--meta-dir', dest='meta_dir', help="Pasta dos arquivos Meta (opcional)", default=META_DIR)
    parser.add_argument('--rescue-nao', dest='rescue_nao', help="Arquivo CSV de não encontrados a ser resgatado (padrão: Relatorio_NaoEncontrados.csv)", default=None)
    parser.add_argument('--fuzzy-threshold', dest='fuzzy_threshold', help="Threshold para fuzzy email quando resgatando (float 0-1)", type=float, default=0.85)
    args = parser.parse_args()

    print("🚀 INICIANDO VALIDAÇÃO DE COBERTURA DE DADOS COMPLETA (com matching avançado)...")
    print(f"📂 Diretório Base: {BASE_DIR}")

    # 1) HubSpot
    print("\n[1/4] 🔍 Localizando arquivo do HubSpot...")
    hub_file = find_hubspot_file(DATA_DIR, preferred_name=args.hub_file)
    if hub_file is None:
        print("❌ Arquivo do HubSpot não encontrado em Data/. Use --hub-file para especificar o nome exato ou substring.")
        sys.exit(1)

    print(f"ℹ️ Arquivo do HubSpot selecionado: {os.path.basename(hub_file)}")

    try:
        df_hs = read_hubspot(hub_file)
    except Exception as e:
        print(f"❌ Erro ao ler HubSpot: {e}")
        sys.exit(1)

    total_hubspot = len(df_hs)
    print(f"✅ HubSpot Carregado: {total_hubspot} leads únicos.")

    # 2) Meta
    print("\n[2/4] 📥 Carregando base Meta (nbfo*)...")
    try:
        df_meta = read_meta_all(args.meta_dir)
        # mostra quantos arquivos nbfo foram lidos para diagnóstico
        files_read = [f for f in os.listdir(args.meta_dir) if f.lower().startswith('nbfo')]
        print(f"✅ Meta Carregado: {len(df_meta)} registros com data de nascimento. Arquivos lidos: {len(files_read)}")
        if len(files_read) <= 20:
            print("   -> Detalhe arquivos:")
            for f in files_read:
                print(f"      - {f}")
    except Exception as e:
        print(f"❌ Erro ao processar Meta: {e}")
        sys.exit(1)

    # Se o usuário pediu resgate específico dos não encontrados, executa aqui e sai
    if args.rescue_nao is not None:
        rescue_file = args.rescue_nao or os.path.join(BASE_DIR, 'Relatorio_NaoEncontrados.csv')
        if not os.path.isfile(rescue_file):
            print(f"❌ Arquivo de não-encontrados não existe: {rescue_file}")
            sys.exit(1)
        print(f"\n🔁 Executando resgate para a lista: {os.path.basename(rescue_file)} (fuzzy_threshold={args.fuzzy_threshold})")
        df_rescue = rescue_nao_encontrados(rescue_file, df_meta, fuzzy_threshold=args.fuzzy_threshold)
        arquivo_rescue = os.path.join(BASE_DIR, 'Relatorio_Resgates_NaoEncontrados.csv')
        df_rescue.to_csv(arquivo_rescue, index=False, sep=';', encoding='utf-8-sig')
        print(f"💾 Resgates salvos em: {arquivo_rescue} (total encontrados: {int(df_rescue['matched'].sum())})")
        # resumo top
        print(df_rescue.head(20).to_string(index=False))

        # Também atualiza a versão preenchida (cria cópia do HubSpot e preenche 'idade' e 'data_nascimento' quando houver resgate por phone/email)
        df_hs_local = df_hs.copy()
        # aplica resgates como preenchimentos (prioridade: email_exact -> phone_exact -> fuzzy)
        resc_map = df_rescue[df_rescue['matched']].copy()
        for _, r in resc_map.iterrows():
            email = r['hub_email']
            idade = r.get('inferred_age', None)
            meta_dob = r.get('meta_dob', None)
            if email in df_hs_local['key_email'].values:
                if meta_dob and pd.notna(meta_dob):
                    df_hs_local.loc[df_hs_local['key_email'] == email, 'data_nascimento_meta'] = meta_dob
                    try:
                        df_hs_local.loc[df_hs_local['key_email'] == email, 'idade'] = compute_age_from_dob(pd.to_datetime(meta_dob))
                    except Exception:
                        pass
                elif idade and not pd.isna(idade):
                    df_hs_local.loc[df_hs_local['key_email'] == email, 'idade'] = idade
                df_hs_local.loc[df_hs_local['key_email'] == email, 'filled_by'] = r.get('method')

        # salva cópia preenchida do HubSpot (CSV)
        base_name = os.path.splitext(os.path.basename(hub_file))[0]
        csv_preenchido = os.path.join(BASE_DIR, f"{base_name}_preenchido_resgates.csv")
        df_hs_local.to_csv(csv_preenchido, index=False, sep=';', encoding='utf-8-sig')
        print(f"💾 HubSpot preenchido salvo em: {csv_preenchido}")

        # ------------------ Export adicional: Excel com aba IDENTICA + aba Campanha ------------------
        try:
            excel_out = os.path.join(BASE_DIR, f"{base_name}_resgatado.xlsx")
            print(f"\n📦 Gerando planilha resgatada: {os.path.basename(excel_out)}")
            # lê arquivo original como df_original para manter estrutura
            if hub_file.lower().endswith(('.xls', '.xlsx')):
                df_original = pd.read_excel(hub_file)
            else:
                df_original = try_read_csv_with_fallback(hub_file) or pd.DataFrame()

            # prepara df_identico (copia de df_original com colunas de data/idade adicionadas)
            df_identico = df_original.copy()
            # encontra coluna de email original para casar (reutiliza detecção)
            original_cols = list(df_identico.columns)
            col_email_orig = next((c for c in original_cols if any(k in str(c).lower() for k in EMAIL_COL_KEYWORDS)), None)
            if col_email_orig is None:
                raise RuntimeError('Coluna de e-mail não encontrada na planilha original.')

            # normaliza emails para mapear
            df_identico['key_email'] = df_identico[col_email_orig].apply(clean_email)
            # adiciona colunas desejadas
            df_identico['data_nascimento_meta'] = ''
            df_identico['idade'] = ''
            df_identico['filled_by'] = ''

            for _, r in resc_map.iterrows():
                email = r['hub_email']
                meta_dob = r.get('meta_dob', None)
                idade = r.get('inferred_age', None)
                method = r.get('method')
                if email in df_identico['key_email'].values:
                    if meta_dob and pd.notna(meta_dob):
                        df_identico.loc[df_identico['key_email'] == email, 'data_nascimento_meta'] = meta_dob
                        try:
                            df_identico.loc[df_identico['key_email'] == email, 'idade'] = compute_age_from_dob(pd.to_datetime(meta_dob))
                        except Exception:
                            pass
                    elif idade and not pd.isna(idade):
                        df_identico.loc[df_identico['key_email'] == email, 'idade'] = idade
                    df_identico.loc[df_identico['key_email'] == email, 'filled_by'] = method

            # prepara aba campanha: pega apenas rescates matched e inclui campaign/source_file
            df_camp = df_rescue[df_rescue['matched']].copy()
            # garantir colunas úteis
            for c in ['source_file', 'campaign']:
                if c not in df_camp.columns:
                    df_camp[c] = ''

            with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                # aba IDENTICA ao Diogo (mas com preenchimentos)
                # remove coluna auxiliar
                df_identico.drop(columns=['key_email'], inplace=True)
                df_identico.to_excel(writer, sheet_name='Diogo_Preenchido', index=False)
                # aba Campanha
                df_camp.to_excel(writer, sheet_name='Campanha', index=False)
            print(f"💾 Planilha resgatada salva em: {excel_out}")
        except Exception as e:
            print(f"⚠️ Falha ao gerar planilha resgatada: {e}")

        sys.exit(0)

    # 3) Cruzamento (LEFT JOIN)
    print("\n[3/4] 🔄 Realizando Cruzamento (De-Para por e-mail exato)...")
    df_final = pd.merge(
        df_hs,
        df_meta,
        on='key_email',
        how='left'
    )

    # Considera ENCONTRADO apenas se data_nascimento_meta não for NaN e não vazia
    df_final['data_nascimento_meta'] = df_final['data_nascimento_meta'].replace(['', 'nan', 'NaN'], np.nan)
    df_final['Status'] = np.where(df_final['data_nascimento_meta'].notna(), 'ENCONTRADO', 'NÃO ENCONTRADO')

    # 3b) Matching avançado para não encontrados (diagnóstico)
    print("\n🔎 Executando matching avançado para diagnóstico (telefone + fuzzy)...")
    df_poss = find_possible_matches(df_hs, df_meta, top_n=5)

    # ------------------ Preenchimento exato (email já feito, agora telefone exato) ------------------
    print("\n🔧 Preenchendo datas por correspondência exata (email/telefone)...")
    df_final, df_changes = apply_exact_fill(df_final, df_meta)

    # [NOVO] Lógica robusta de identificação e sincronização de colunas de data
    all_cols = df_final.columns.tolist()
    
    # 1. Identificar colunas alvo com flexibilidade
    # Prioridade 1: 'data de nascimento do aluno'
    col_aluno = next((c for c in all_cols if 'data de nascimento do aluno' in str(c).lower()), None)
    # Prioridade 2: 'data de nascimento' (sem candidato)
    if not col_aluno:
        col_aluno = next((c for c in all_cols if 'data de nascimento' in str(c).lower() and 'candidato' not in str(c).lower()), None)
    # Prioridade 3: 'nascimento' (sem candidato)
    if not col_aluno:
        col_aluno = next((c for c in all_cols if 'nascimento' in str(c).lower() and 'candidato' not in str(c).lower()), None)
        
    col_candidato = next((c for c in all_cols if 'data de nascimento do candidato' in str(c).lower()), None)
    
    # 2. Atualizar Coluna do Aluno com dados da Meta (se encontrados)
    mask_meta = df_final['data_nascimento_meta'].notna() & (df_final['data_nascimento_meta'] != '')
    
    if col_aluno:
        df_final.loc[mask_meta, col_aluno] = df_final.loc[mask_meta, 'data_nascimento_meta']
    else:
        # Se não achou coluna do aluno, cria uma padrão
        col_aluno = 'Data de Nascimento do Aluno'
        df_final[col_aluno] = df_final['data_nascimento_meta']

    # 3. Sincronizar Candidato = Aluno
    # Se coluna candidato não existe, define nome padrão
    if not col_candidato:
        col_candidato = 'Data de Nascimento do Candidato'
    
    # Remove "Nenhum Valor" ou valores vazios da coluna candidato
    df_final[col_candidato] = df_final[col_candidato].replace(['Nenhum Valor', '', 'nan', 'NaN'], np.nan)
    
    # Copia integralmente TUDO de Aluno para Candidato (valores ou vazio/NaN)
    df_final[col_candidato] = df_final[col_aluno]

    # Relatório de mudanças
    if not df_changes.empty:
        arquivo_mud = os.path.join(BASE_DIR, f"Relatorio_Mudancas_Preenchimento.csv")
        df_changes.to_csv(arquivo_mud, index=False, sep=';', encoding='utf-8-sig')
        print(f"💾 Relatório de mudanças salvo em: {arquivo_mud} ({len(df_changes)} alterações)")
    else:
        print("ℹ️ Nenhuma alteração por preenchimento exato encontrada.")

    # Atualiza contagem marcada por telefone (se houver)
    phone_exact = df_changes[df_changes['method'] == 'phone_exact'] if not df_changes.empty else pd.DataFrame()
    if not phone_exact.empty:
        matched_emails = phone_exact['hub_email'].unique().tolist()
        df_final.loc[df_final['key_email'].isin(matched_emails), 'Status'] = 'ENCONTRADO (phone)'

    # 4) Output
    print("\n[4/4] 📊 Resultados:")
    total_encontrados = int(((df_final['Status'] == 'ENCONTRADO') | (df_final['Status'] == 'ENCONTRADO (phone)')).sum())
    cobertura = (total_encontrados / total_hubspot) * 100 if total_hubspot else 0.0

    print("-" * 60)
    print(f"Total Leads HubSpot: {total_hubspot}")
    print(f"Total Encontrados no Meta (inclui phone matches): {total_encontrados}")
    print(f"📈 Cobertura: {cobertura:.2f}%")
    print("-" * 60)

    print("\n🔎 Amostra das primeiras 10 linhas (E-mail + Status):")
    sample = df_final[['key_email', 'Status']].rename(columns={'key_email': 'E-mail'})
    print(sample.head(10).to_string(index=False))

    arquivo_saida = os.path.join(BASE_DIR, 'Relatorio_Cobertura.csv')
    df_final.to_csv(arquivo_saida, index=False, sep=';', encoding='utf-8-sig')
    print(f"\n💾 Relatório principal salvo em: {arquivo_saida}")

    # Salvar não encontrados e possíveis matches para análise
    nao_encontrados = df_final[df_final['Status'].str.contains('NÃO ENCONTRADO')].copy()
    arquivo_nao = os.path.join(BASE_DIR, 'Relatorio_NaoEncontrados.csv')
    nao_encontrados[['key_email', 'Status', 'phone_hs']].rename(columns={'key_email': 'E-mail'}).to_csv(arquivo_nao, index=False, sep=';', encoding='utf-8-sig')
    print(f"💾 Relatório de não encontrados salvo em: {arquivo_nao}")

    if not df_poss.empty:
        # salva versão limpa e completa
        arquivo_poss = os.path.join(BASE_DIR, 'Relatorio_Possiveis_Matches.csv')
        df_poss.to_csv(arquivo_poss, index=False, sep=';', encoding='utf-8-sig')
        print(f"💾 Relatório de possíveis matches (limpo) salvo em: {arquivo_poss}")

        # salva também apenas os top candidates por hub_email (rank == 1)
        df_poss_top = df_poss[df_poss['rank'] == 1].reset_index(drop=True)
        arquivo_poss_top = os.path.join(BASE_DIR, 'Relatorio_Possiveis_Matches_Top.csv')
        df_poss_top.to_csv(arquivo_poss_top, index=False, sep=';', encoding='utf-8-sig')
        print(f"💾 Relatório de possíveis matches (top) salvo em: {arquivo_poss_top}")

        # estatísticas rápidas para diagnóstico
        print(f"ℹ️ Possíveis matches: {len(df_poss)} registros ({df_poss['hub_email'].nunique()} leads únicos com candidatos). Top candidates: {len(df_poss_top)}")
    else:
        print("ℹ️ Nenhum possível match gerado pelo algoritmo de fuzzy/telefone.")

    # ------------------ Export: Versão do arquivo do Diogo com aba de descobertas ------------------
    try:
        base_name = os.path.splitext(os.path.basename(hub_file))[0]
        excel_out = os.path.join(BASE_DIR, f"{base_name}_com_descobertas.xlsx")
        print(f"\n📦 Gerando versão do arquivo com aba de descobertas: {os.path.basename(excel_out)}")
        # lê o arquivo original sem alterar estrutura para colocar na primeira aba
        if hub_file.lower().endswith(('.xls', '.xlsx')):
            df_original = pd.read_excel(hub_file)
        else:
            df_original = try_read_csv_with_fallback(hub_file) or pd.DataFrame()

        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
            df_original.to_excel(writer, sheet_name='Original', index=False)
            # Aba preenchida (mesma estrutura + coluna 'idade' e 'filled_by')
            # Prepara df_preenchido mantendo colunas originais na mesma ordem
            df_pre = df_final.copy()
            # garante colunas preservadas e adiciona 'idade' e 'filled_by' ao final
            original_cols = df_original.columns.tolist()
            for c in ['idade', 'filled_by', 'meta_matched_email', 'meta_matched_phone']:
                if c not in df_pre.columns:
                    df_pre[c] = ''
            
            # Garante que colunas de data (Aluno/Candidato) estejam na saída, mesmo se criadas agora
            extra_cols = ['idade', 'filled_by', 'meta_matched_email', 'meta_matched_phone']
            # Adiciona colunas de data se não estiverem na lista original
            for col_date in ['Data de Nascimento do Aluno', 'Data de Nascimento do Candidato']:
                if col_date in df_pre.columns and col_date not in original_cols:
                    extra_cols.append(col_date)

            # reordena: coloca original cols primeiro (se existirem), depois as col novas
            cols_final = [c for c in original_cols if c in df_pre.columns] + [c for c in extra_cols if c in df_pre.columns]
            
            # SINCRONIZAR COLUNAS DE DATA: G = H
            # Identifica colunas de data
            col_aluno_final = next((c for c in cols_final if 'data de nascimento do aluno' in str(c).lower()), None)
            col_candidato_final = next((c for c in cols_final if 'data de nascimento do candidato' in str(c).lower()), None)
            
            # Se ambas colunas existem, sincroniza: coloca em candidato exatamente o que está em aluno
            if col_aluno_final and col_candidato_final:
                # Remove "Nenhum valor" e variações de AMBAS as colunas
                df_pre[col_aluno_final] = df_pre[col_aluno_final].replace(['(Nenhum valor)', 'Nenhum Valor', 'Nenhum valor', '', 'nan', 'NaN'], np.nan)
                df_pre[col_candidato_final] = df_pre[col_candidato_final].replace(['(Nenhum valor)', 'Nenhum Valor', 'Nenhum valor', '', 'nan', 'NaN'], np.nan)
                # Replica coluna aluno para candidato (sincroniza totalmente)
                df_pre[col_candidato_final] = df_pre[col_aluno_final]
            
            df_pre[cols_final].to_excel(writer, sheet_name='Preenchido', index=False)

            # Aba com mudanças aplicadas automaticamente
            if not df_changes.empty:
                df_changes.to_excel(writer, sheet_name='Descobertas', index=False)
            else:
                # escreve um dataframe vazio com cabeçalho padrão
                pd.DataFrame(columns=['hub_email','hub_phone','old_field','old_value','new_field','new_value','method','meta_email','meta_phone']).to_excel(writer, sheet_name='Descobertas', index=False)
            # Se houver possíveis matches, escreve em aba separada (full) e a aba Top
            if not df_poss.empty:
                df_poss.to_excel(writer, sheet_name='Possiveis_Matches', index=False)
                df_poss_top.to_excel(writer, sheet_name='Possiveis_Matches_Top', index=False)
        print(f"💾 Versão com descobertas salva em: {excel_out}")
    except Exception as e:
        print(f"⚠️ Falha ao gerar versão com descobertas: {e}")


if __name__ == '__main__':
    main()
