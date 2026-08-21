import json
import os
import snowflake.connector

COLUMN_MAP = {
    'REQUIREMENT_ID':    'Requirement ID',
    'REQUIREMENT_NAME':  'Requirement Name',
    'REQUIREMENT_URL':   'Requirement URL',
    'FEATURE_SPEC':      'Feature Spec',
    'RELEASED_IN':       'Released in',
    'REQMT_FOR_VERSION': 'Reqmt for Version',
    'TEST_CASE_ID':      'Test Case ID',
    'TEST_CASE_NAME':    'Test Case Name',
    'TEST_CASE_STATE':   'Test Case State',
    'RELEVANT_PRODUCT':  'Relevant Product',
    'TEST_CASE_RELEASE': 'Test Case Release',
    'TEST_CASE_URL':     'Test Case URL',
    'RISK_ID':           'Risk ID',
    'RISK_NAME':         'Risk Name',
    'RISK_URL':          'Risk URL',
    'TEST_POINT_ID':     'Test Point ID',
    'TEST_PLAN_ID':      'Test Plan ID',
    'TEST_SUITE_ID':     'Test Suite ID',
    'TEST_SUITE_NAME':   'Test Suite Name',
    'TEST_PLAN_URL':     'Test Plan URL',
    'TEST_SUITE_URL':    'Test Suite URL',
    'TEST_RESULT':       'Test Result',
    'TEST_RESULT_DATE':  'Test Result Date',
    'TEST_PLAN_TAGS':    'Test Plan Tags',
    'STICR_ID':          'STICR ID',
    'STICR_NAME':        'STICR Name',
    'STICR_URL':         'STICR URL',
    'STICR_STATE':       'STICR State',
    'STICR_RELEASE':     'STICR Release',
}

conn = snowflake.connector.connect(
    account=os.environ['SNOWFLAKE_ACCOUNT'],
    user=os.environ['SNOWFLAKE_USER'],
    authenticator='programmatic_access_token',
    token=os.environ['SNOWFLAKE_PAT'],
    warehouse='SNOWFLAKE_LEARNING_WH',
    database='PHILIPS_APPS',
    schema='IX_TOOLS_HUB',
)

cur = conn.cursor()
cur.execute('SELECT * FROM TRACEABILITY_MATRIX')
cols = [COLUMN_MAP[d[0]] for d in cur.description]

rows = []
for row in cur:
    obj = {col: '' if val is None else str(val) for col, val in zip(cols, row)}
    rows.append(obj)

cur.close()
conn.close()

with open('data.json', 'w', encoding='utf-8') as f:
    json.dump(rows, f)

print(f'Wrote {len(rows)} rows to data.json')
