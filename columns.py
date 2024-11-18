

SOL_COLUMN_MAPPING = {
    'projectId': 'External Project ID',
    'projectName': 'Content Name',
    'target':'Target Language',
    'tgroup':'Vendor (Translator Group)',
    'step': 'Step',
    'sol_total':'Billing Quantity',
    'sol_total_unit': 'Unit',
    'sls_unit':'SLS Unit Key',
}

SOL_UNIT_NAME_MAPPING = {
    'sol_no_match': 'NO_MATCH',
    'sol_100_match': '100_MATCH',
    'sol_repetitions':'REPETITIONS',
    'sol_95':'95-99_MATCH',
    'sol_85':'85-94_MATCH',
    'sol_75':'75-84_MATCH',
    'sol_non_trans':'NON_TRANSLATABLE',
    'sol_ice':'ICE_MATCH',
    'sol_mt': 'MT',
    'sol_edc1':'EDC1',
    'sol_edc2':'EDC2',
    'sol_edc3':'EDC3',
    'sol_edc4':'EDC4',
    'sol_edc5':'EDC5',
    'sol_edc6':'EDC6',
    'sol_edc7':'EDC7',
}

SOL_GROUPBY_COLUMNS = [ 'projectId', 'projectName', 'target', 'tgroup']
SOL_SUM_COLUMNS = ['sol_total'] + list(SOL_UNIT_NAME_MAPPING.keys())

TEW_COLUMN_MAPPING = {
    # First three columns have the same name, preparing for a merge.
    'projectId': 'XTM_PROJECT_ID',
    'target':'JOB_TARGET_LANGUAGE',
    'tgroup':'LSP_NAME',
    'tew_name':'XTM_PROJECT_NAME',
    'tew_total':'STATS_VOLUME_TOTAL_WORDS',
    'tew_no_match':'STATS_VOLUME_NO_MATCHING_WORDS',
    'tew_100_match':'STATS_VOLUME_LEVERAGED_WORDS',
    'tew_repetitions':'STATS_VOLUME_REPEATS_WORDS',
    'tew_95_a': 'STATS_VOLUME_HIGH_FUZZY_MATCH_WORDS',
    'tew_95_b': 'STATS_VOLUME_HIGH_FUZZY_REPEATS_WORDS',
    'tew_85_a': 'STATS_VOLUME_MEDIUM_FUZZY_MATCH_WORDS',
    'tew_85_b': 'STATS_VOLUME_MEDIUM_FUZZY_REPEATS_WORDS',
    'tew_75_a': 'STATS_VOLUME_LOW_FUZZY_MATCH_WORDS',
    'tew_75_b': 'STATS_VOLUME_LOW_FUZZY_REPEATS_WORDS',
    'tew_non_trans': 'STATS_VOLUME_NON_TRANSLATABLE_WORDS',
    'tew_ice':'STATS_VOLUME_ICE_MATCH_WORDS',
    'tew_mt':'STATS_VOLUME_MACHINE_TRANSLATION_WORDS',
    'tew_edc1':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY1_WORDCOUNT',
    'tew_edc2':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY2_WORDCOUNT',
    'tew_edc3':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY3_WORDCOUNT',
    'tew_edc4':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY4_WORDCOUNT',
    'tew_edc5':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY5_WORDCOUNT',
    'tew_edc6':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY6_WORDCOUNT',
    'tew_edc7':'STATS_VOLUME_MACHINE_TRANSLATION_EDC_WORDS_CATEGORY7_WORDCOUNT',
}

TEW_GROUPBY_COLUMNS = [ 'projectId','tew_name', 'target', 'tgroup' ]

JOIN_KEY_COLUMNS = [ 'projectId', 'target', 'tgroup' ]

FINAL_COLUMN_ORDER = [
    # Group-By Columns
    'projectId', 'projectName', 'target', 'tgroup',
    # Generated as 'ok' if totals match, else 'mismatch'
    'status',
    # START OF COLUMN PAIRS.
    'sol_total', 'tew_total',
    'sol_no_match', 'tew_no_match',
    'sol_100_match','tew_100_match',
    'sol_repetitions','tew_repetitions',
    'sol_95','tew_95',
    'sol_85','tew_85',
    'sol_75','tew_75',
    'sol_non_trans','tew_non_trans',
    'sol_ice','tew_ice',
    'sol_mt','tew_mt',
    'sol_edc1','tew_edc1',
    'sol_edc2','tew_edc2',
    'sol_edc3','tew_edc3',
    'sol_edc4','tew_edc4',
    'sol_edc5','tew_edc5',
    'sol_edc6','tew_edc6',
    'sol_edc7','tew_edc7'
]

FINAL_COLUMN_PAIRS = [
    ('sol_no_match', 'tew_no_match'),
    ('sol_100_match','tew_100_match'),
    ('sol_repetitions','tew_repetitions'),
    ('sol_95','tew_95'),
    ('sol_85','tew_85'),
    ('sol_75','tew_75'),
    ('sol_non_trans','tew_non_trans'),
    ('sol_ice','tew_ice'),
    # These two columns are not related.
    #    ('sol_mt','tew_mt'),
    ('sol_edc1','tew_edc1'),
    ('sol_edc2','tew_edc2'),
    ('sol_edc3','tew_edc3'),
    ('sol_edc4','tew_edc4'),
    ('sol_edc5','tew_edc5'),
    ('sol_edc6','tew_edc6'),
    ('sol_edc7','tew_edc7')
]