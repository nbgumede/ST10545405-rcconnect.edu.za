from docx import Document
from docx.shared import Inches,Pt,RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import pandas as pd, json
OUT='/mnt/data'
# source inventory
sources=pd.DataFrame([
 {'source':'South African National Cancer Registry (NCR/NHLS)','scope':'National pathology-based cancer surveillance','data_noted':'Demographic, clinical and reporting-source data; annual statistics by sex, age and population groups','public_access':'Annual reports publicly listed; row-level access requires permission','classification_readiness':'Not confirmed as public patient-level extract','url':'https://www.nicd.ac.za/centres/national-cancer-registry/'},
 {'source':'NCR Cancer Statistics','scope':'South African cancer statistics','data_noted':'Research sources include public hospitals, private facilities, SA Oncology Consortium, hospice, paediatric oncology, CHOC and NHLS','public_access':'Published statistics; permission contact listed by NCR','classification_readiness':'Requires approved row-level extract and outcome definition','url':'https://www.nicd.ac.za/centres/national-cancer-registry/cancer-statistics/'},
 {'source':'CMJAH/NHLS breast cancer molecular classification research','scope':'Core-biopsy breast cancers diagnosed at Charlotte Maxeke Johannesburg Academic Hospital during 2019','data_noted':'Age, tumour grade, histological subtype, ER, PR, HER2 and Ki67; 265 cases described','public_access':'Research report located; raw row-level file not established as publicly downloadable','classification_readiness':'Strong candidate if de-identified records are authorised','url':'https://wiredspace.wits.ac.za/bitstreams/11648622-6f66-4d43-92ad-e3001957e6c0/download'}
])
sources.to_csv(f'{OUT}/SA_Cancer_Data_Source_Inventory.csv',index=False)
# data request template
schema=pd.DataFrame([
 ('record_id','Anonymous unique identifier','ID only; exclude from model'),('age_at_diagnosis','Age at confirmed diagnosis','Numeric'),('diagnosis_date','Date of confirmed diagnosis','Date'),('tumour_grade','Histological tumour grade','Ordinal'),('histological_subtype','Pathology subtype','Categorical'),('er_status','Estrogen receptor result','Binary/categorical'),('pr_status','Progesterone receptor result','Binary/categorical'),('her2_status','HER2 result','Categorical'),('ki67_index','Ki67 proliferative index','Numeric'),('molecular_subtype','Luminal A/B, HER2-enriched or basal/triple-negative','Candidate classification target'),('benefit_eligible','Eligibility after authorised scheme adjudication','Separate target only if governed and available')
],columns=['field','definition','modelling_role'])
schema.to_csv(f'{OUT}/SA_Cancer_Requested_Data_Schema.csv',index=False)
# notebook template
cells=[]
def md(s):cells.append({'cell_type':'markdown','metadata':{},'source':[x+'\n' for x in s.splitlines()]})
def code(s):cells.append({'cell_type':'code','execution_count':None,'metadata':{},'outputs':[],'source':[x+'\n' for x in s.splitlines()]})
md('# South African Cancer Classification Analysis\n## Local-data implementation notebook\nThis notebook is intentionally not populated with synthetic patient records. It becomes executable after an authorised, de-identified NCR/NHLS/CMJAH row-level extract is supplied.')
md('## Expected research target\nPreferred target: `molecular_subtype` for a multiclass breast-cancer research model. `benefit_eligible` must not be used unless it comes from an authorised scheme adjudication dataset with documented policy labels.')
code("""import pandas as pd, numpy as np, matplotlib.pyplot as plt, seaborn as sns
from pathlib import Path
DATA_PATH = Path('authorised_sa_breast_cancer_extract.csv')
if not DATA_PATH.exists():
    raise FileNotFoundError('Add the authorised de-identified South African extract. Do not substitute synthetic patient data.')
df = pd.read_csv(DATA_PATH)
print(df.shape)
df.head()""")
md('## Data-quality and EDA checks')
code("""print(df.info())
print(df.isna().mean().sort_values(ascending=False))
print('Duplicates:', df.duplicated().sum())
print(df.describe(include='all').T)
sns.countplot(data=df, x='molecular_subtype'); plt.xticks(rotation=30); plt.show()""")
md('## Leakage-safe preprocessing, feature selection and model tuning')
code("""from sklearn.model_selection import train_test_split, StratifiedKFold, GridSearchCV
from sklearn.compose import ColumnTransformer
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import OneHotEncoder, StandardScaler
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report, ConfusionMatrixDisplay, roc_auc_score

features=['age_at_diagnosis','tumour_grade','histological_subtype','er_status','pr_status','her2_status','ki67_index']
target='molecular_subtype'
X=df[features].copy(); y=df[target].copy()
num=['age_at_diagnosis','ki67_index']; cat=[c for c in features if c not in num]
pre=ColumnTransformer([
 ('num',Pipeline([('impute',SimpleImputer(strategy='median')),('scale',StandardScaler())]),num),
 ('cat',Pipeline([('impute',SimpleImputer(strategy='most_frequent')),('onehot',OneHotEncoder(handle_unknown='ignore'))]),cat)
])
pipe=Pipeline([('pre',pre),('model',LogisticRegression(max_iter=5000,solver='saga'))])
X_train,X_test,y_train,y_test=train_test_split(X,y,test_size=.2,stratify=y,random_state=42)
params={'model__C':[0.01,0.1,1,10], 'model__penalty':['l1','l2'], 'model__class_weight':[None,'balanced']}
cv=StratifiedKFold(5,shuffle=True,random_state=42)
search=GridSearchCV(pipe,params,scoring='f1_macro',cv=cv,n_jobs=-1)
search.fit(X_train,y_train)
pred=search.predict(X_test)
print(search.best_params_)
print(classification_report(y_test,pred))
ConfusionMatrixDisplay.from_predictions(y_test,pred,xticks_rotation=30); plt.show()""")
md('## Governance checkpoint\nDo not present molecular-subtype prediction as benefit approval. Eligibility requires scheme rules, verified clinical evidence, authorised adjudication and human review. Record data provenance, approval, feature definitions, exclusions and model version.')
nb={'cells':cells,'metadata':{'kernelspec':{'display_name':'Python 3','language':'python','name':'python3'},'language_info':{'name':'python','version':'3'}},'nbformat':4,'nbformat_minor':5}
with open(f'{OUT}/SA_Cancer_Classification_Analysis.ipynb','w') as f:json.dump(nb,f,indent=2)
# report
D=Document(); s=D.sections[0]; s.top_margin=Inches(.7);s.bottom_margin=Inches(.7);s.left_margin=Inches(.75);s.right_margin=Inches(.75)
D.styles['Normal'].font.name='Aptos';D.styles['Normal'].font.size=Pt(10)
for st in ['Heading 1','Heading 2']:
 D.styles[st].font.name='Aptos Display';D.styles[st].font.color.rgb=RGBColor(0,92,65)
p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;r=p.add_run('South African Cancer Classification Data Plan');r.bold=True;r.font.size=Pt(20);r.font.color.rgb=RGBColor(0,92,65)
p=D.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.add_run('Local replacement for the international breast-cancer dataset').italic=True
D.add_heading('Executive summary',1)
D.add_paragraph('The recommended local source is the South African National Cancer Registry within the National Health Laboratory Service. The NCR is suitable for South African contextual relevance because it is the country’s main source of national cancer incidence data and conducts pathology-based surveillance. A closely aligned Johannesburg study used NHLS histology records from Charlotte Maxeke Johannesburg Academic Hospital and described 265 breast-cancer core-biopsy cases with age, tumour grade, histological subtype, ER, PR, HER2 and Ki67 variables. However, the raw patient-level study file was not established as publicly downloadable. Consequently, this pack does not fabricate records or report model performance. It provides the same report/notebook support structure, a source inventory and the required data schema for an authorised extract.')
D.add_heading('1. Local datasource selection and justification',1)
D.add_paragraph('Primary source: South African National Cancer Registry (NCR), National Health Laboratory Service (NHLS).')
for x in ['Quality: the NCR receives pathology-confirmed cancer information from public and private laboratories and reports national cancer statistics.','Suitability: pathology, demographic and clinical variables are relevant to cancer classification research and are locally grounded.','Research alignment: the CMJAH/NHLS study demonstrates a viable breast-cancer feature set for molecular subtype classification.','Access limitation: published annual statistics and a research report are not equivalent to a downloadable row-level training table. Patient-level data require permission, de-identification and governance approval.']:
 D.add_paragraph(x,style='List Bullet')
D.add_heading('2. Revised classification question',1)
D.add_paragraph('Recommended research question: Can age, tumour grade, histological subtype and immunohistochemical markers (ER, PR, HER2 and Ki67) classify breast-cancer molecular subtype in an authorised South African NHLS-derived dataset?')
D.add_paragraph('This formulation matches the variables explicitly described in the Johannesburg study. It is preferable to claiming direct benefit eligibility because the located research data do not contain medical-scheme policy decisions.')
D.add_heading('3. Algorithm choice',1)
D.add_paragraph('Multinomial logistic regression is proposed as the transparent baseline because molecular subtype has multiple categories and the expected features contain both numeric and categorical variables. One-hot encoding, scaling, regularisation and class weighting can be implemented in a leakage-safe pipeline. Macro F1 is the proposed model-selection metric because it gives equal importance to each subtype, while per-class recall, precision, confusion matrix and one-vs-rest ROC-AUC should be reported where probabilities and class counts permit.')
D.add_heading('4. Succinct analysis plan',1)
t=D.add_table(rows=1,cols=3);t.style='Light Shading Accent 1';t.alignment=WD_TABLE_ALIGNMENT.CENTER
for i,h in enumerate(['Stage','Action','Evidence']):t.rows[0].cells[i].text=h
for row in [('Acquire','Obtain ethics/data-owner approval and a de-identified row-level NCR/NHLS/CMJAH extract.','Approval, data dictionary and provenance.'),('EDA','Check completeness, duplicates, coding, class balance, implausible values and distributions.','Quality log and plots.'),('Features','Use age, grade, histology, ER, PR, HER2 and Ki67 where available; exclude identifiers and post-outcome fields.','Leakage review and retained-feature list.'),('Train','Stratified split; preprocessing pipeline; tune regularisation and class weights with stratified CV.','Best parameters and cross-validation results.'),('Evaluate','Macro F1, per-class recall/precision, confusion matrix and ROC-AUC where valid.','Held-out results with uncertainty and error analysis.'),('Report','Explain provenance, cleaning, model, limitations, governance and non-eligibility boundary.','Client report and reproducible notebook.')]:
 c=t.add_row().cells
 for i,v in enumerate(row):c[i].text=v
D.add_heading('5. Conducting and evaluating the study',1)
D.add_paragraph('The accompanying notebook contains commented Python for loading the authorised extract, EDA, missing-data treatment, categorical encoding, scaling, stratified splitting, hyperparameter tuning and evaluation. Execution intentionally stops when the approved file is absent. This prevents accidental use of invented patient data and avoids reporting unsupported accuracy results.')
D.add_heading('6. Required data fields',1)
t=D.add_table(rows=1,cols=3);t.style='Light Shading Accent 1'
for i,h in enumerate(schema.columns):t.rows[0].cells[i].text=h.replace('_',' ').title()
for _,r in schema.iterrows():
 c=t.add_row().cells
 for i,v in enumerate(r):c[i].text=str(v)
D.add_heading('7. Medical-scheme use limitation',1)
D.add_paragraph('A molecular-subtype classifier can support research or case prioritisation, but it cannot independently adjudicate a dreaded-disease benefit. A production eligibility model would require a separate, governed scheme dataset containing verified policy labels and applicable clinical evidence. Final decisions should remain with authorised clinical and claims personnel.')
D.add_heading('References',1)
D.add_paragraph('National Institute for Communicable Diseases. National Cancer Registry. Available at: https://www.nicd.ac.za/centres/national-cancer-registry/ (Accessed 20 October 2026).')
D.add_paragraph('National Institute for Communicable Diseases. Cancer Statistics. Available at: https://www.nicd.ac.za/centres/national-cancer-registry/cancer-statistics/ (Accessed 20 October 2026).')
D.add_paragraph('Breast cancer: molecular classification using immunohistochemical expression at the NHLS, Charlotte Maxeke Johannesburg Academic Hospital, Johannesburg. WIReDSpace research record.')
D.save(f'{OUT}/SA_Cancer_Classification_Report.docx')
print('created')
