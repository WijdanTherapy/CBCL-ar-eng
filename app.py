import streamlit as st
from groq import Groq
import io, os, smtplib, re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import date
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, Image as RLImage, HRFlowable,
                                 PageBreak, KeepTogether)

# ══════════════════════════════════════════════════════════════
#  CONFIG
# ══════════════════════════════════════════════════════════════
GMAIL_USER      = "Wijdan.psyc@gmail.com"
GMAIL_PASS      = "rias eeul lyuu stce"
RECIPIENT_EMAIL = "Wijdan.psyc@gmail.com"
LOGO_FILE       = "logo.png"

PDF_DARK   = colors.HexColor('#1C1917')
PDF_WARM   = colors.HexColor('#8B7355')
PDF_CREAM  = colors.HexColor('#F7F3EE')
PDF_BORDER = colors.HexColor('#DDD5C8')
PDF_RED    = colors.HexColor('#C62828')
PDF_ORANGE = colors.HexColor('#F57C00')
PDF_YELLOW = colors.HexColor('#FBC02D')
PDF_GREEN  = colors.HexColor('#388E3C')
PDF_HEADER = colors.HexColor('#2D2926')

# ══════════════════════════════════════════════════════════════
#  ITEMS  — full 113 items (120 response slots including 56a-h)
# ══════════════════════════════════════════════════════════════
# Keys: integers 1-55, strings "56a"-"56h", integers 57-113
ITEM_KEYS = (list(range(1, 56)) +
             ["56a","56b","56c","56d","56e","56f","56g","56h"] +
             list(range(57, 114)))

ITEMS_EN = {
    1:"Acts too young for his/her age",
    2:"Drinks alcohol without parents' approval",
    3:"Argues a lot",
    4:"Fails to finish things he/she starts",
    5:"There is very little he/she enjoys",
    6:"Bowel movements outside toilet",
    7:"Bragging, boasting",
    8:"Can't concentrate, can't pay attention for long",
    9:"Can't get his/her mind off certain thoughts; obsessions",
    10:"Can't sit still, restless, or hyperactive",
    11:"Clings to adults or too dependent",
    12:"Complains of loneliness",
    13:"Confused or seems to be in a fog",
    14:"Cries a lot",
    15:"Cruel to animals",
    16:"Cruelty, bullying, or meanness to others",
    17:"Daydreams or gets lost in his/her thoughts",
    18:"Deliberately harms self or attempts suicide",
    19:"Demands a lot of attention",
    20:"Destroys his/her own things",
    21:"Destroys things belonging to his/her family or others",
    22:"Disobedient at home",
    23:"Disobedient at school",
    24:"Doesn't eat well",
    25:"Doesn't get along with other kids",
    26:"Doesn't seem to feel guilty after misbehaving",
    27:"Easily jealous",
    28:"Breaks rules at home, school, or elsewhere",
    29:"Fears certain animals, situations, or places other than school",
    30:"Fears going to school",
    31:"Fears he/she might think or do something bad",
    32:"Feels he/she has to be perfect",
    33:"Feels or complains that no one loves him/her",
    34:"Feels others are out to get him/her",
    35:"Feels worthless or inferior",
    36:"Gets hurt a lot, accident-prone",
    37:"Gets in many fights",
    38:"Gets teased a lot",
    39:"Hangs around with others who get in trouble",
    40:"Hears sounds or voices that aren't there",
    41:"Impulsive or acts without thinking",
    42:"Would rather be alone than with others",
    43:"Lying or cheating",
    44:"Bites fingernails",
    45:"Nervous, highstrung, or tense",
    46:"Nervous movements or twitching",
    47:"Nightmares",
    48:"Not liked by other kids",
    49:"Constipated, doesn't move bowels",
    50:"Too fearful or anxious",
    51:"Feels dizzy or lightheaded",
    52:"Feels too guilty",
    53:"Overeating",
    54:"Overtired without good reason",
    55:"Overweight",
    "56a":"56a. Physical problems — aches or pains (not stomach or headaches)",
    "56b":"56b. Physical problems — headaches",
    "56c":"56c. Physical problems — nausea, feels sick",
    "56d":"56d. Physical problems — problems with eyes (not corrected by glasses)",
    "56e":"56e. Physical problems — rashes or other skin problems",
    "56f":"56f. Physical problems — stomachaches",
    "56g":"56g. Physical problems — vomiting, throwing up",
    "56h":"56h. Physical problems — other physical problems",
    57:"Physically attacks people",
    58:"Picks nose, skin, or other parts of body",
    59:"Plays with own sex parts in public",
    60:"Plays with own sex parts too much",
    61:"Poor school work",
    62:"Poorly coordinated or clumsy",
    63:"Prefers being with older kids",
    64:"Prefers being with younger kids",
    65:"Refuses to talk",
    66:"Repeats certain acts over and over; compulsions",
    67:"Runs away from home",
    68:"Screams a lot",
    69:"Secretive, keeps things to self",
    70:"Sees things that aren't there",
    71:"Self-conscious or easily embarrassed",
    72:"Sets fires",
    73:"Sexual problems",
    74:"Showing off or clowning",
    75:"Too shy or timid",
    76:"Sleeps less than most kids",
    77:"Sleeps more than most kids during day and/or night",
    78:"Inattentive or easily distracted",
    79:"Speech problem",
    80:"Stares blankly",
    81:"Steals at home",
    82:"Steals outside the home",
    83:"Stores up too many things he/she doesn't need",
    84:"Strange behavior",
    85:"Strange ideas",
    86:"Stubborn, sullen, or irritable",
    87:"Sudden changes in mood or feelings",
    88:"Sulks a lot",
    89:"Suspicious",
    90:"Swearing or obscene language",
    91:"Talks about killing self",
    92:"Talks or walks in sleep",
    93:"Talks too much",
    94:"Teases a lot",
    95:"Temper tantrums or hot temper",
    96:"Thinks about sex too much",
    97:"Threatens people",
    98:"Thumb-sucking",
    99:"Smokes, chews, or sniffs tobacco",
    100:"Trouble sleeping",
    101:"Truancy, skips school",
    102:"Underactive, slow moving, or lacks energy",
    103:"Unhappy, sad, or depressed",
    104:"Unusually loud",
    105:"Uses drugs for nonmedical purposes (not alcohol or tobacco)",
    106:"Vandalism",
    107:"Wets self during the day",
    108:"Wets the bed",
    109:"Whining",
    110:"Wishes to be of opposite sex",
    111:"Withdrawn, doesn't get involved with others",
    112:"Worries",
    113:"Other problems not listed above",
}

ITEMS_AR = {
    1:"يتصرف بشكل أصغر من سنه",
    2:"يشرب الكحول دون علم والديه أو موافقتهما",
    3:"كثير الجدال",
    4:"لا يُتمّ ما يبدأه",
    5:"نادراً ما يجد شيئاً يستمتع به",
    6:"يتبرز خارج المرحاض",
    7:"يتفاخر ويتباهى بشكل مفرط",
    8:"لا يستطيع التركيز أو الانتباه لفترة طويلة",
    9:"لا يستطيع إخراج أفكار بعينها من ذهنه — أفكار وسواسية",
    10:"لا يستطيع الجلوس بهدوء، مضطرب أو مفرط النشاط",
    11:"يتعلق بالكبار ويعتمد عليهم بشكل مفرط",
    12:"يشكو من الوحدة والشعور بالعزلة",
    13:"يبدو مرتبكاً أو غير واعٍ لما يدور حوله",
    14:"كثير البكاء",
    15:"قاسٍ في تعامله مع الحيوانات",
    16:"يتسلط على الآخرين أو يمارس التنمر ضدهم",
    17:"يحلم أحلام يقظة أو يغرق في أفكاره",
    18:"يؤذي نفسه عمداً أو حاول إيذاء نفسه",
    19:"يطالب باهتمام مستمر ومفرط من الآخرين",
    20:"يُتلف أغراضه الشخصية",
    21:"يُتلف ممتلكات أفراد الأسرة أو الآخرين",
    22:"عاصٍ وغير مطيع في المنزل",
    23:"عاصٍ وغير مطيع في المدرسة",
    24:"شهيته ضعيفة أو لا يأكل بشكل طبيعي",
    25:"علاقاته مع الأطفال الآخرين متوترة أو سيئة",
    26:"لا يشعر بتأنيب الضمير أو الذنب بعد الإساءة",
    27:"سريع الغيرة",
    28:"يخالف القواعد في المنزل والمدرسة وغيرها",
    29:"يخاف من حيوانات أو مواقف أو أماكن بعينها غير المدرسة",
    30:"يخاف من الذهاب إلى المدرسة",
    31:"يخشى أن يُقدم على أفعال أو يراودَه أفكار سيئة",
    32:"يشعر أنه يجب أن يكون مثالياً في كل شيء",
    33:"يشعر أو يشتكي من أن لا أحد يحبه",
    34:"يشعر أن الآخرين يتآمرون عليه أو يريدون إيذاءه",
    35:"يشعر بعدم قيمته أو أنه أقل من الآخرين",
    36:"يتعرض للإصابات والحوادث بصورة متكررة",
    37:"يدخل في شجارات كثيرة",
    38:"يتعرض للسخرية والمضايقة من أقرانه بشكل متكرر",
    39:"يصاحب أقراناً يُسببون المشاكل",
    40:"يسمع أصواتاً أو أشخاصاً غير موجودين في الواقع",
    41:"متهور ويتصرف دون تفكير أو تروٍّ",
    42:"يُفضّل البقاء وحيداً بدلاً من مصاحبة الآخرين",
    43:"يكذب أو يغش",
    44:"يعض أظافره",
    45:"متوتر أو مضطرب بصورة واضحة",
    46:"حركات عصبية أو ارتعاش لا إرادي",
    47:"يُعاني من كوابيس",
    48:"لا يحبه أقرانه",
    49:"يعاني من إمساك وصعوبة في التبرز",
    50:"يعاني من خوف مفرط أو قلق شديد",
    51:"يشعر بالدوار أو بخفة مفاجئة في الرأس",
    52:"يشعر بذنب مفرط وغير مبرر",
    53:"يُفرط في الأكل",
    54:"يشعر بتعب وإرهاق شديد دون سبب واضح",
    55:"يعاني من زيادة ملحوظة في الوزن",
    "56a":"٥٦أ. أوجاع وآلام جسدية غير محددة (غير آلام المعدة أو الصداع)",
    "56b":"٥٦ب. صداع متكرر",
    "56c":"٥٦ج. غثيان أو شعور بالمرض",
    "56d":"٥٦د. مشكلات في البصر غير المُعالجة بالنظارة",
    "56e":"٥٦هـ. طفح جلدي أو مشكلات جلدية",
    "56f":"٥٦و. آلام في المعدة",
    "56g":"٥٦ز. تقيؤ أو إفراز",
    "56h":"٥٦ح. مشكلات جسدية أخرى",
    57:"يهاجم الآخرين جسدياً",
    58:"ينقر أنفه أو جلده أو أجزاء من جسمه باستمرار",
    59:"يلمس أعضاءه الجنسية في الأماكن العامة",
    60:"يلمس أعضاءه الجنسية بشكل مفرط",
    61:"أداؤه الدراسي ضعيف",
    62:"ضعيف التناسق الحركي أو أخرق",
    63:"يُفضّل مصاحبة من هم أكبر منه سناً",
    64:"يُفضّل مصاحبة من هم أصغر منه سناً",
    65:"يرفض الكلام",
    66:"يُكرر أفعالاً بعينها مراراً — سلوك قهري",
    67:"يهرب من المنزل",
    68:"يصرخ كثيراً",
    69:"كتوم ويحتفظ بكل شيء لنفسه",
    70:"يرى أشياء غير موجودة في الواقع",
    71:"يشعر بالإحراج بسهولة أو يتأثر بسرعة",
    72:"يُشعل الحرائق",
    73:"لديه مشكلات جنسية",
    74:"يُكثر من التهريج وإثبات الذات بشكل مبالغ فيه",
    75:"خجول جداً أو مُتردد",
    76:"ينام أقل من معظم أقرانه",
    77:"ينام أكثر من معظم أقرانه نهاراً أو ليلاً",
    78:"سهل التشتت وضعيف التركيز",
    79:"لديه مشكلة في الكلام أو النطق",
    80:"ينظر في الفراغ بنظرة فارغة",
    81:"يسرق في المنزل",
    82:"يسرق خارج المنزل",
    83:"يكتنز أشياء لا يحتاج إليها",
    84:"تصرفات غريبة",
    85:"أفكار غريبة",
    86:"عنيد أو متجهم أو سريع الانفعال",
    87:"مزاجه يتغير فجأة وبشكل ملحوظ",
    88:"يتعكر مزاجه ويُكثر من التذمر",
    89:"مريب أو شكّاك في الآخرين",
    90:"يستخدم ألفاظاً نابية أو بذيئة",
    91:"يتحدث عن إيذاء نفسه أو الانتحار",
    92:"يتكلم أو يمشي أثناء نومه",
    93:"يتكلم كثيراً",
    94:"يُكثر من إزعاج الآخرين والتحرش بهم",
    95:"نوبات غضب حادة أو انفجارات انفعالية",
    96:"يُكثر من التفكير في الأمور الجنسية",
    97:"يُهدد الآخرين",
    98:"يمص إبهامه",
    99:"يُدخن أو يمضغ أو يشم التبغ",
    100:"صعوبة في النوم",
    101:"يتغيب عن المدرسة بدون إذن",
    102:"خامل وبطيء الحركة أو يفتقر إلى الطاقة",
    103:"يشعر بالتعاسة أو الحزن أو الاكتئاب",
    104:"صوته عالٍ بشكل غير معتاد",
    105:"يتعاطى مواد أو عقاقير لأغراض غير طبية",
    106:"يُخرّب ممتلكات الآخرين",
    107:"يبلل ملابسه أو نفسه أثناء النهار",
    108:"يبلل فراشه أثناء النوم",
    109:"كثير التذمر والشكوى",
    110:"يتمنى لو كان من الجنس الآخر",
    111:"مُنسحب ولا يُشارك الآخرين في أي نشاط",
    112:"كثير القلق",
    113:"مشكلات أخرى لم تُذكر أعلاه",
}

# ══════════════════════════════════════════════════════════════
#  SUBSCALES
# ══════════════════════════════════════════════════════════════
SUBSCALES = {
    "ANX":  {"name_en":"Anxious/Depressed",       "name_ar":"القلق / الاكتئاب",
             "items":[14,29,30,31,32,33,35,45,50,52,71,91,112],"color":"#6A1B9A","group":"INT"},
    "WIT":  {"name_en":"Withdrawn/Depressed",      "name_ar":"الانسحاب / الاكتئاب",
             "items":[5,42,65,69,75,102,103,111],"color":"#1565C0","group":"INT"},
    "SOM":  {"name_en":"Somatic Complaints",       "name_ar":"الشكاوى الجسدية",
             "items":[47,49,51,54,"56a","56b","56c","56f","56g"],"color":"#4E342E","group":"INT"},
    "SOC":  {"name_en":"Social Problems",          "name_ar":"المشكلات الاجتماعية",
             "items":[11,12,25,27,34,36,38,48,62,64,79],"color":"#0277BD","group":"MIX"},
    "THO":  {"name_en":"Thought Problems",         "name_ar":"مشكلات التفكير",
             "items":[9,18,40,46,58,59,60,66,70,76,83,84,85,92,100],"color":"#827717","group":"MIX"},
    "ATT":  {"name_en":"Attention Problems",       "name_ar":"مشكلات الانتباه",
             "items":[1,4,8,10,13,17,41,61,78],"color":"#E65100","group":"MIX"},
    "RUL":  {"name_en":"Rule-Breaking Behavior",   "name_ar":"السلوك المخالف للقواعد",
             "items":[2,26,28,39,43,63,67,72,73,81,82,90,96,99,101,105,106],"color":"#B71C1C","group":"EXT"},
    "AGG":  {"name_en":"Aggressive Behavior",      "name_ar":"السلوك العدواني",
             "items":[3,16,19,20,21,22,23,37,57,68,86,87,88,89,94,95,97,104],"color":"#C62828","group":"EXT"},
    "INT":  {"name_en":"Internalizing Problems",   "name_ar":"المشكلات الانطوائية","items":[],"color":"#4527A0","group":"BROAD"},
    "EXT":  {"name_en":"Externalizing Problems",   "name_ar":"المشكلات السلوكية الخارجية","items":[],"color":"#BF360C","group":"BROAD"},
    "TOT":  {"name_en":"Total Problems",           "name_ar":"إجمالي المشكلات","items":[],"color":"#1B5E20","group":"BROAD"},
    "DEP":  {"name_en":"DSM: Depressive",          "name_ar":"الاكتئاب (DSM)",
             "items":[5,14,18,24,35,52,54,76,77,91,100,102,103],"color":"#283593","group":"DSM"},
    "ANX_D":{"name_en":"DSM: Anxiety",             "name_ar":"القلق (DSM)",
             "items":[11,29,30,31,45,47,50,71,112],"color":"#6A1B9A","group":"DSM"},
    "SOM_D":{"name_en":"DSM: Somatic",             "name_ar":"الجسدي (DSM)",
             "items":["56a","56b","56c","56f","56g"],"color":"#4E342E","group":"DSM"},
    "ADHD_D":{"name_en":"DSM: ADHD",              "name_ar":"ADHD (DSM)",
              "items":[4,8,10,41,78,93,104],"color":"#E65100","group":"DSM"},
    "ODD":  {"name_en":"DSM: Oppositional Defiant","name_ar":"التمرد والعصيان (DSM)",
             "items":[3,22,23,86,95],"color":"#AD1457","group":"DSM"},
    "CON":  {"name_en":"DSM: Conduct",             "name_ar":"اضطراب السلوك (DSM)",
             "items":[15,16,21,26,28,37,39,43,57,67,72,81,82,90,97,101,106],"color":"#BF360C","group":"DSM"},
}
SUBSCALES["INT"]["items"] = list(set(SUBSCALES["ANX"]["items"]+SUBSCALES["WIT"]["items"]+SUBSCALES["SOM"]["items"]))
SUBSCALES["EXT"]["items"] = list(set(SUBSCALES["RUL"]["items"]+SUBSCALES["AGG"]["items"]))
SUBSCALES["TOT"]["items"] = list(set(
    SUBSCALES["ANX"]["items"]+SUBSCALES["WIT"]["items"]+SUBSCALES["SOM"]["items"]+
    SUBSCALES["SOC"]["items"]+SUBSCALES["THO"]["items"]+SUBSCALES["ATT"]["items"]+
    SUBSCALES["RUL"]["items"]+SUBSCALES["AGG"]["items"]
))

NORMS = {
    "ANX":(3.5,3.8),"WIT":(2.2,2.5),"SOM":(1.8,2.4),"SOC":(2.5,2.9),
    "THO":(1.2,2.1),"ATT":(4.0,3.5),"RUL":(1.8,2.6),"AGG":(5.2,5.5),
    "INT":(7.5,6.8),"EXT":(7.0,7.5),"TOT":(24.0,18.0),
    "DEP":(2.8,3.2),"ANX_D":(1.6,2.2),"SOM_D":(0.8,1.5),
    "ADHD_D":(2.5,2.8),"ODD":(2.0,2.5),"CON":(1.2,2.0),
}
SCALE_ORDER = ["ANX","WIT","SOM","SOC","THO","ATT","RUL","AGG","INT","EXT","TOT","DEP","ANX_D","SOM_D","ADHD_D","ODD","CON"]
CRITICAL_KEYS = [18,40,72,91,105,107,108]

def get_level_en(t):
    return "Clinical Range" if t>=70 else "Borderline Clinical" if t>=65 else "Worth Monitoring" if t>=60 else "Normal Range"
def get_level_ar(t):
    return "النطاق الإكلينيكي" if t>=70 else "حدي الإكلينيكي" if t>=65 else "يستحق المتابعة" if t>=60 else "ضمن الطبيعي"
def get_bar_color(t):
    return "#D32F2F" if t>=70 else "#F57C00" if t>=65 else "#FBC02D" if t>=60 else "#388E3C"

def raw_to_t(raw, key):
    mean, sd = NORMS[key]
    return max(20, min(90, round(50+10*(raw-mean)/sd))) if sd else 50

def compute_scores(responses):
    results = {}
    for key, info in SUBSCALES.items():
        raw = sum(responses.get(k,0) for k in info["items"])
        results[key] = {"raw":raw,"t":raw_to_t(raw,key),"max_raw":len(info["items"])*2}
    return results

# ══════════════════════════════════════════════════════════════
#  CHARTS
# ══════════════════════════════════════════════════════════════
def _chart_buf(fig):
    buf = io.BytesIO(); plt.savefig(buf,format='png',dpi=150,bbox_inches='tight')
    plt.close(fig); buf.seek(0); return buf.read()

def make_syndrome_chart(scores):
    keys = ["ANX","WIT","SOM","SOC","THO","ATT","RUL","AGG"]
    t_vals = [scores[k]["t"] for k in keys]
    fig,ax = plt.subplots(figsize=(11,6)); fig.patch.set_facecolor('#F7F3EE'); ax.set_facecolor('#F7F3EE')
    y = np.arange(len(keys))
    bars = ax.barh(y,t_vals,color=[get_bar_color(t) for t in t_vals],edgecolor='white',lw=0.8,height=0.6)
    for xv,lbl,col in [(60,'T=60','#FBC02D'),(65,'T=65','#F57C00'),(70,'T=70','#D32F2F')]:
        ax.axvline(x=xv,color=col,linestyle='--',lw=1.2,alpha=0.8,label=lbl)
    ax.axvspan(70,97,alpha=0.07,color='#D32F2F'); ax.axvspan(65,70,alpha=0.06,color='#F57C00')
    for bar_,val in zip(bars,t_vals):
        ax.text(bar_.get_width()+0.5,bar_.get_y()+bar_.get_height()/2,str(val),
                va='center',ha='left',fontsize=9,fontweight='bold',color='#1C1917')
    ax.set_yticks(y); ax.set_yticklabels([SUBSCALES[k]["name_en"] for k in keys],fontsize=10.5)
    ax.set_xlim(20,97); ax.set_xlabel('T-Score',fontsize=11,fontweight='bold',color='#1C1917')
    ax.set_title("CBCL/6-18 — Syndrome Scale T-Score Profile",fontsize=13,fontweight='bold',color='#1C1917',pad=12)
    ax.legend(loc='lower right',fontsize=8.5,framealpha=0.7)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.grid(axis='x',linestyle=':',alpha=0.4); plt.tight_layout(); return _chart_buf(fig)

def make_broadband_chart(scores):
    keys = ["INT","EXT","TOT"]; labels = ["Internalizing","Externalizing","Total Problems"]
    t_vals = [scores[k]["t"] for k in keys]; raw_v = [scores[k]["raw"] for k in keys]
    fig,ax = plt.subplots(figsize=(7,4.5)); fig.patch.set_facecolor('#F7F3EE'); ax.set_facecolor('#F7F3EE')
    x = np.arange(len(keys))
    bars = ax.bar(x,t_vals,color=[get_bar_color(t) for t in t_vals],edgecolor='white',lw=1.2,width=0.55)
    for xv,lbl,col in [(65,'T=65','#F57C00'),(70,'T=70','#D32F2F')]:
        ax.axhline(y=xv,color=col,linestyle='--',lw=1.3,alpha=0.85,label=lbl)
    for bar_,t,raw in zip(bars,t_vals,raw_v):
        ax.text(bar_.get_x()+bar_.get_width()/2,bar_.get_height()+0.8,
                f"T={t}\n(Raw={raw})",ha='center',va='bottom',fontsize=9,fontweight='bold',color='#1C1917')
    ax.set_xticks(x); ax.set_xticklabels(labels,fontsize=11)
    ax.set_ylim(20,95); ax.set_ylabel('T-Score',fontsize=10)
    ax.set_title("Broadband Problem Scales",fontsize=12,fontweight='bold',color='#1C1917',pad=10)
    ax.legend(fontsize=9,framealpha=0.7)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.grid(axis='y',linestyle=':',alpha=0.4); plt.tight_layout(); return _chart_buf(fig)

def make_dsm_chart(scores):
    keys = ["DEP","ANX_D","SOM_D","ADHD_D","ODD","CON"]
    labels = ["Depressive","Anxiety","Somatic","ADHD","ODD","Conduct"]
    t_vals = [scores[k]["t"] for k in keys]
    fig,ax = plt.subplots(figsize=(10,5)); fig.patch.set_facecolor('#F7F3EE'); ax.set_facecolor('#F7F3EE')
    x = np.arange(len(keys))
    bars = ax.bar(x,t_vals,color=[get_bar_color(t) for t in t_vals],edgecolor='white',lw=1.0,width=0.6)
    for xv,lbl,col in [(65,'T=65 Borderline','#F57C00'),(70,'T=70 Clinical','#D32F2F')]:
        ax.axhline(y=xv,color=col,linestyle='--',lw=1.3,alpha=0.85,label=lbl)
    for bar_,val in zip(bars,t_vals):
        ax.text(bar_.get_x()+bar_.get_width()/2,bar_.get_height()+0.6,str(val),
                ha='center',va='bottom',fontsize=10,fontweight='bold',color='#1C1917')
    ax.set_xticks(x); ax.set_xticklabels(labels,fontsize=11)
    ax.set_ylim(20,95); ax.set_ylabel('T-Score',fontsize=10)
    ax.set_title("CBCL/6-18 — DSM-Oriented Scale Scores",fontsize=12,fontweight='bold',color='#1C1917',pad=10)
    ax.legend(fontsize=9,framealpha=0.7)
    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
    ax.grid(axis='y',linestyle=':',alpha=0.4); plt.tight_layout(); return _chart_buf(fig)

def make_pie_chart(responses):
    counts = [sum(1 for v in responses.values() if v==i) for i in range(3)]
    fig,ax = plt.subplots(figsize=(6,4.5)); fig.patch.set_facecolor('#F7F3EE')
    wedges,texts,autotexts = ax.pie(counts,
        labels=['0 — Not True','1 — Somewhat True','2 — Very True'],
        colors=['#388E3C','#FBC02D','#D32F2F'],autopct='%1.0f%%',startangle=90,
        wedgeprops={'edgecolor':'white','linewidth':1.5})
    for at in autotexts: at.set_fontsize(9); at.set_fontweight('bold')
    ax.set_title('Response Distribution',fontsize=11,fontweight='bold',color='#1C1917')
    plt.tight_layout(); return _chart_buf(fig)

# ══════════════════════════════════════════════════════════════
#  GROQ
# ══════════════════════════════════════════════════════════════
def generate_report_en(child_name, age, gender, rater, relationship, scores, responses):
    elevated_syn = [SUBSCALES[k]['name_en'] for k in ["ANX","WIT","SOM","SOC","THO","ATT","RUL","AGG"] if scores[k]["t"]>=65]
    elevated_dsm = [SUBSCALES[k]['name_en'] for k in ["DEP","ANX_D","SOM_D","ADHD_D","ODD","CON"]     if scores[k]["t"]>=65]
    critical_flagged = [str(k) for k in CRITICAL_KEYS if responses.get(k,0)>=1]
    int_t=scores["INT"]["t"]; ext_t=scores["EXT"]["t"]; tot_t=scores["TOT"]["t"]
    score_block = "\n".join(
        ["\n  SYNDROME SCALES:"] +
        [f"    {SUBSCALES[k]['name_en']}: Raw={scores[k]['raw']}, T={scores[k]['t']} — {get_level_en(scores[k]['t'])}"
         for k in ["ANX","WIT","SOM","SOC","THO","ATT","RUL","AGG"]] +
        ["\n  BROADBAND SCALES:"] +
        [f"    {SUBSCALES[k]['name_en']}: Raw={scores[k]['raw']}, T={scores[k]['t']} — {get_level_en(scores[k]['t'])}"
         for k in ["INT","EXT","TOT"]] +
        ["\n  DSM-ORIENTED SCALES:"] +
        [f"    {SUBSCALES[k]['name_en']}: Raw={scores[k]['raw']}, T={scores[k]['t']} — {get_level_en(scores[k]['t'])}"
         for k in ["DEP","ANX_D","SOM_D","ADHD_D","ODD","CON"]]
    )
    prompt = f"""You are a licensed child clinical psychologist writing a professional CBCL/6-18 report.

CHILD: {child_name} | AGE: {age} | GENDER: {gender}
RATER: {rater} ({relationship or 'Parent/Caregiver'})
ASSESSMENT: Child Behavior Checklist for Ages 6-18 (CBCL/6-18) — Achenbach & Rescorla (2001)
DATE: {date.today().strftime('%B %d, %Y')} | RATING PERIOD: Past 6 months

T-SCORES (T≥70=Clinical; T=65-69=Borderline Clinical; T<65=Normal):
{score_block}

ELEVATED SYNDROME SCALES (T≥65): {', '.join(elevated_syn) if elevated_syn else 'None'}
ELEVATED DSM SCALES (T≥65): {', '.join(elevated_dsm) if elevated_dsm else 'None'}
CRITICAL ITEMS ENDORSED (score≥1): {', '.join(critical_flagged) if critical_flagged else 'None'}

RULES: No formal diagnoses. Findings as clinical hypotheses only.
No markdown (**, ##, ---). Section titles: ALL CAPS numbered.
Reference specific T-scores throughout.

WRITE THIS COMPLETE REPORT:

CHILD BEHAVIOR CHECKLIST (CBCL/6-18) — CLINICAL REPORT
Child | {child_name}
Age | {age}  ·  Gender | {gender}
Rater | {rater} ({relationship or 'Parent/Caregiver'})
Date | {date.today().strftime('%B %d, %Y')}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

CLINICAL SUMMARY
5-7 sentences: Internalizing T={int_t}, Externalizing T={ext_t}, Total T={tot_t}. Most elevated scales. Key clinical concerns.

1. REFERRAL AND ASSESSMENT OVERVIEW
CBCL/6-18 description, ASEBA system, purpose, who completed it, rating period.

2. INTERNALIZING PROBLEMS
Internalizing T={int_t}. Deep analysis of Anxious/Depressed (T={scores['ANX']['t']}), Withdrawn/Depressed (T={scores['WIT']['t']}), Somatic Complaints (T={scores['SOM']['t']}). Behavioral correlates, daily and school impact.

3. EXTERNALIZING PROBLEMS
Externalizing T={ext_t}. Rule-Breaking Behavior (T={scores['RUL']['t']}), Aggressive Behavior (T={scores['AGG']['t']}). School, family, and social implications.

4. MIXED-DOMAIN SYNDROME SCALES
Social Problems (T={scores['SOC']['t']}), Thought Problems (T={scores['THO']['t']}), Attention Problems (T={scores['ATT']['t']}). Clinical significance of each.

5. DSM-ORIENTED SCALE ANALYSIS
For each scale: T-score, range, diagnostic framework implications as hypotheses only.

6. TOTAL PROBLEMS AND OVERALL SEVERITY
Total T={tot_t}. Overall behavioral/emotional burden. Internalizing vs Externalizing dominance.

7. CRITICAL ITEMS AND SAFETY CONCERNS
Address endorsed critical items (self-harm, suicidal ideation, fire-setting, hallucinations, substance use).
If none: "No critical items were endorsed at score ≥1."

8. STRENGTHS AND PROTECTIVE FACTORS
Scales in normal range as relative strengths.

9. INTEGRATED CLINICAL IMPRESSIONS
Full profile synthesis. Primary diagnostic hypotheses. Differential considerations.

10. RECOMMENDATIONS
12 specific evidence-based recommendations:
a) Mental health intervention priorities
b) School-based accommodations
c) Parenting and family-based interventions
d) Further psychometric or diagnostic evaluation
e) Medical consultation if indicated
f) Community and social supports

11. SUMMARY FOR CLINICAL RECORDS
One paragraph: "The CBCL/6-18 was completed by {rater} for {child_name} (age {age}, {gender}). Results indicate..."
"""
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
    r = client.chat.completions.create(
        model="llama-3.3-70b-versatile",
        messages=[{"role":"user","content":prompt}],
        max_tokens=4000
    )
    return r.choices[0].message.content.strip()

# ══════════════════════════════════════════════════════════════
#  PDF BUILDER
# ══════════════════════════════════════════════════════════════
def _t_band(t):
    if t>=70: return PDF_RED,"Clinical Range"
    if t>=65: return PDF_ORANGE,"Borderline Clinical"
    if t>=60: return PDF_YELLOW,"Worth Monitoring"
    return PDF_GREEN,"Normal Range"

def _styles():
    s={}
    s['title']  = ParagraphStyle('t',fontName='Helvetica-Bold',fontSize=16,textColor=PDF_DARK,spaceAfter=4,alignment=TA_CENTER)
    s['sub']    = ParagraphStyle('sub',fontName='Helvetica',fontSize=9,textColor=PDF_WARM,spaceAfter=2,alignment=TA_CENTER)
    s['sec']    = ParagraphStyle('sec',fontName='Helvetica-Bold',fontSize=11,textColor=PDF_WARM,spaceBefore=14,spaceAfter=4)
    s['body']   = ParagraphStyle('body',fontName='Helvetica',fontSize=9.5,textColor=PDF_DARK,leading=14,spaceAfter=5)
    s['small']  = ParagraphStyle('small',fontName='Helvetica',fontSize=8,textColor=PDF_WARM,leading=11)
    return s

def build_pdf(report_text, scores, charts, child_name, age, gender, rater, relationship, responses):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf,pagesize=A4,leftMargin=2*cm,rightMargin=2*cm,topMargin=2*cm,bottomMargin=2*cm)
    S = _styles(); W = A4[0]-4*cm; story = []

    if os.path.exists(LOGO_FILE):
        try:
            logo=RLImage(LOGO_FILE,width=5*cm,height=2.2*cm); logo.hAlign='CENTER'
            story.append(logo); story.append(Spacer(1,4))
        except: pass

    story.append(Paragraph("Child Behavior Checklist — Clinical Report",S['title']))
    story.append(Paragraph("CBCL/6-18  ·  Achenbach & Rescorla (2001)  ·  ASEBA",S['sub']))
    story.append(HRFlowable(width=W,thickness=1,color=PDF_BORDER,spaceAfter=10))

    rel_str = relationship or "Parent/Caregiver"
    demo = [
        [Paragraph('<b>Child</b>',S['small']),Paragraph(child_name or '—',S['body']),
         Paragraph('<b>Age</b>',S['small']),Paragraph(str(age) or '—',S['body'])],
        [Paragraph('<b>Gender</b>',S['small']),Paragraph(gender or '—',S['body']),
         Paragraph('<b>Rater</b>',S['small']),Paragraph(f"{rater} ({rel_str})",S['body'])],
        [Paragraph('<b>Date</b>',S['small']),Paragraph(date.today().strftime('%B %d, %Y'),S['body']),
         Paragraph('<b>Instrument</b>',S['small']),Paragraph('CBCL/6-18 · 113 items · 0–1–2 scale',S['body'])],
        [Paragraph('<b>Rating period</b>',S['small']),Paragraph('Past 6 months',S['body']),
         Paragraph('<b>Cutoffs</b>',S['small']),Paragraph('T≥65 Borderline · T≥70 Clinical',S['body'])],
    ]
    dt=Table(demo,colWidths=[2.5*cm,6.2*cm,2.5*cm,6.2*cm])
    dt.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,-1),PDF_CREAM),('BOX',(0,0),(-1,-1),0.5,PDF_BORDER),
        ('INNERGRID',(0,0),(-1,-1),0.3,PDF_BORDER),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),('LEFTPADDING',(0,0),(-1,-1),6)]))
    story.append(KeepTogether([dt])); story.append(Spacer(1,10))

    def add_score_tbl(title,keys):
        rows=[[Paragraph('<b>Scale</b>',S['small']),Paragraph('<b>Raw</b>',S['small']),
               Paragraph('<b>T-Score</b>',S['small']),Paragraph('<b>Range</b>',S['small'])]]
        ts=[('BACKGROUND',(0,0),(-1,0),PDF_HEADER),('TEXTCOLOR',(0,0),(-1,0),colors.white),
            ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),8.5),
            ('BOX',(0,0),(-1,-1),0.5,PDF_BORDER),('INNERGRID',(0,0),(-1,-1),0.3,PDF_BORDER),
            ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
            ('LEFTPADDING',(0,0),(-1,-1),5),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,PDF_CREAM])]
        for ri,k in enumerate(keys,start=1):
            s=scores[k]; t=s['t']; col_,lbl_=_t_band(t)
            ts+=[('BACKGROUND',(2,ri),(2,ri),col_),('TEXTCOLOR',(2,ri),(2,ri),colors.black),
                 ('FONTNAME',(2,ri),(2,ri),'Helvetica-Bold')]
            rows.append([Paragraph(SUBSCALES[k]['name_en'],S['body']),Paragraph(str(s['raw']),S['body']),
                         Paragraph(str(t),S['body']),Paragraph(lbl_,S['body'])])
        tbl=Table(rows,colWidths=[6.5*cm,1.8*cm,1.8*cm,7.3*cm]); tbl.setStyle(TableStyle(ts))
        story.append(KeepTogether([Paragraph(title,S['sec']),tbl,Spacer(1,6)]))

    add_score_tbl("SYNDROME SCALE SCORES",["ANX","WIT","SOM","SOC","THO","ATT","RUL","AGG"])
    add_score_tbl("BROADBAND SCALE SCORES",["INT","EXT","TOT"])
    add_score_tbl("DSM-ORIENTED SCALE SCORES",["DEP","ANX_D","SOM_D","ADHD_D","ODD","CON"])

    if charts.get("syndrome"):
        story.append(Paragraph("SYNDROME SCALE PROFILE",S['sec']))
        story.append(RLImage(io.BytesIO(charts["syndrome"]),width=W,height=9*cm)); story.append(Spacer(1,6))
    if charts.get("broadband") and charts.get("dsm"):
        bb=RLImage(io.BytesIO(charts["broadband"]),width=W*0.42,height=7*cm)
        dm=RLImage(io.BytesIO(charts["dsm"]),width=W*0.56,height=7*cm)
        two=Table([[bb,dm]],colWidths=[W*0.44,W*0.56])
        two.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE'),('LEFTPADDING',(1,0),(1,0),8)]))
        story.append(Paragraph("BROADBAND AND DSM-ORIENTED PROFILES",S['sec']))
        story.append(two); story.append(Spacer(1,6))
    if charts.get("pie"):
        story.append(Paragraph("RESPONSE DISTRIBUTION",S['sec']))
        pi=RLImage(io.BytesIO(charts["pie"]),width=8*cm,height=6*cm); pi.hAlign='LEFT'
        story.append(pi); story.append(Spacer(1,6))

    story.append(HRFlowable(width=W,thickness=0.5,color=PDF_BORDER))
    story.append(Paragraph("CLINICAL NARRATIVE REPORT",S['sec']))
    sec_pat=re.compile(r'^\d+\.\s+[A-Z][A-Z\s&/()\-:]+$')
    hdr_words={"CHILD BEHAVIOR CHECKLIST (CBCL/6-18) — CLINICAL REPORT","CLINICAL SUMMARY"}
    for line in report_text.split('\n'):
        ls=line.strip()
        if not ls: story.append(Spacer(1,4)); continue
        if ls.startswith('━') or ls.startswith('═'):
            story.append(HRFlowable(width=W,thickness=0.4,color=PDF_BORDER,spaceAfter=4)); continue
        is_sec=sec_pat.match(ls) or ls in hdr_words or ls.upper() in hdr_words or ls=="CLINICAL SUMMARY"
        if is_sec: story.append(Paragraph(ls,S['sec'])); continue
        if '|' in ls:
            parts=[p.strip() for p in ls.split('|') if p.strip()]
            if len(parts)>=2 and (parts[0].lower(),parts[1].lower()) not in [("field","value"),("child",""),("scale","")]:
                mini=Table([[Paragraph(p,S['body']) for p in parts]],colWidths=[W/len(parts)]*len(parts))
                mini.setStyle(TableStyle([('BOX',(0,0),(-1,-1),0.3,PDF_BORDER),
                    ('INNERGRID',(0,0),(-1,-1),0.3,PDF_BORDER),('BACKGROUND',(0,0),(-1,-1),PDF_CREAM),
                    ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),('LEFTPADDING',(0,0),(-1,-1),5)]))
                story.append(KeepTogether([mini])); continue
        story.append(Paragraph(ls,S['body']))

    story.append(PageBreak())
    story.append(Paragraph("ITEM RESPONSES — FULL RATING TABLE",S['sec']))
    story.append(Paragraph("0=Not True  ·  1=Somewhat True  ·  2=Very True  ·  ⚠ = Critical Item",S['small']))
    story.append(Spacer(1,6))
    r_labels={0:"0 — Not True",1:"1 — Somewhat True",2:"2 — Very True"}
    r_colors={0:PDF_GREEN,1:PDF_YELLOW,2:PDF_RED}
    rows=[[Paragraph('<b>#</b>',S['small']),Paragraph('<b>Item</b>',S['small']),Paragraph('<b>Rating</b>',S['small'])]]
    ts_cmds=[('BACKGROUND',(0,0),(-1,0),PDF_HEADER),('TEXTCOLOR',(0,0),(-1,0),colors.white),
             ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7.5),
             ('BOX',(0,0),(-1,-1),0.5,PDF_BORDER),('INNERGRID',(0,0),(-1,-1),0.3,PDF_BORDER),
             ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('TOPPADDING',(0,0),(-1,-1),2),
             ('BOTTOMPADDING',(0,0),(-1,-1),2),('LEFTPADDING',(0,0),(-1,-1),4)]
    for i,key in enumerate(ITEM_KEYS):
        val=responses.get(key,0); bg=colors.white if i%2==0 else PDF_CREAM
        is_crit=key in CRITICAL_KEYS
        ts_cmds+=[('BACKGROUND',(0,i+1),(1,i+1),bg),('BACKGROUND',(2,i+1),(2,i+1),r_colors[val]),
                  ('FONTNAME',(2,i+1),(2,i+1),'Helvetica-Bold')]
        if is_crit: ts_cmds+=[('TEXTCOLOR',(0,i+1),(0,i+1),colors.Color(0.8,0,0)),
                               ('FONTNAME',(0,i+1),(0,i+1),'Helvetica-Bold')]
        num_str=f"⚠{key}" if is_crit else str(key)
        rows.append([Paragraph(num_str,S['small']),Paragraph(ITEMS_EN[key],S['small']),
                     Paragraph(r_labels[val],S['small'])])
    item_tbl=Table(rows,colWidths=[1.2*cm,12.3*cm,4*cm],repeatRows=1)
    item_tbl.setStyle(TableStyle(ts_cmds)); story.append(item_tbl)
    story.append(Spacer(1,12))
    story.append(HRFlowable(width=W,thickness=0.5,color=PDF_BORDER)); story.append(Spacer(1,4))
    story.append(Paragraph(
        "Strictly confidential. For qualified mental health professionals only. "
        "T≥65=Borderline Clinical; T≥70=Clinical Range. "
        "© 2001 T.M. Achenbach, ASEBA, University of Vermont.",S['small']))
    doc.build(story); buf.seek(0); return buf

# ══════════════════════════════════════════════════════════════
#  EMAIL
# ══════════════════════════════════════════════════════════════
def send_email(child_name, buf_pdf, fn_pdf, scores):
    date_str=date.today().strftime('%B %d, %Y')
    elevated=[(k,scores[k]["t"]) for k in SCALE_ORDER if scores[k]["t"]>=65]
    elev_html="".join(
        f"<tr><td style='padding:4px 0;color:#6B5B45;'>{SUBSCALES[k]['name_en']}</td>"
        f"<td><strong style='color:#D9534F;'>T={t}</strong></td></tr>" for k,t in elevated
    ) or "<tr><td colspan='2' style='color:#4CAF50;'>No scales elevated ≥ 65</td></tr>"
    msg=MIMEMultipart('mixed'); msg['From']=GMAIL_USER; msg['To']=RECIPIENT_EMAIL
    msg['Subject']=f"[CBCL/6-18] {child_name} — {date_str}"
    body=f"""<html><body style="font-family:Georgia,serif;color:#1C1917;background:#F7F3EE;padding:20px;">
  <div style="max-width:560px;margin:0 auto;background:white;border:1px solid #DDD5C8;border-radius:4px;padding:28px;">
    <h2 style="font-weight:300;font-size:20px;color:#1C1917;">CBCL/6-18 Report</h2>
    <p style="color:#8B7355;font-size:11px;text-transform:uppercase;letter-spacing:.08em;">
      Child Behavior Checklist for Ages 6–18 · Achenbach & Rescorla (2001)</p>
    <hr style="border:none;border-top:1px solid #DDD5C8;margin:16px 0;">
    <table style="width:100%;font-size:13px;border-collapse:collapse;">
      <tr><td style="padding:5px 0;color:#8B7355;width:40%;">Child</td><td><strong>{child_name}</strong></td></tr>
      <tr><td style="color:#8B7355;">Date</td><td>{date_str}</td></tr>
      <tr><td style="color:#8B7355;">Internalizing T</td><td><strong>{scores['INT']['t']}</strong> — {get_level_en(scores['INT']['t'])}</td></tr>
      <tr><td style="color:#8B7355;">Externalizing T</td><td><strong>{scores['EXT']['t']}</strong> — {get_level_en(scores['EXT']['t'])}</td></tr>
      <tr><td style="color:#8B7355;">Total Problems T</td><td><strong>{scores['TOT']['t']}</strong> — {get_level_en(scores['TOT']['t'])}</td></tr>
    </table>
    <hr style="border:none;border-top:1px solid #DDD5C8;margin:16px 0;">
    <p style="font-size:12px;color:#8B7355;font-weight:bold;margin-bottom:6px;">Elevated Scales (T≥65)</p>
    <table style="width:100%;font-size:12px;border-collapse:collapse;">{elev_html}</table>
    <hr style="border:none;border-top:1px solid #DDD5C8;margin:16px 0;">
    <p style="font-size:12px;">📄 English PDF report attached.</p>
    <p style="font-size:10px;color:#8B7355;font-style:italic;">Confidential — for the treating clinician only.</p>
  </div></body></html>"""
    msg.attach(MIMEText(body,'html'))
    buf_pdf.seek(0)
    part=MIMEBase('application','pdf'); part.set_payload(buf_pdf.read()); encoders.encode_base64(part)
    part.add_header('Content-Disposition','attachment',filename=fn_pdf); msg.attach(part)
    with smtplib.SMTP_SSL('smtp.gmail.com',465) as srv:
        srv.login(GMAIL_USER,GMAIL_PASS); srv.sendmail(GMAIL_USER,RECIPIENT_EMAIL,msg.as_string())

# ══════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════
st.set_page_config(page_title="CBCL/6-18 Assessment",page_icon="📋",layout="wide")
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=EB+Garamond:ital,wght@0,400;0,500;0,600;1,400&family=Cairo:wght@400;600&display=swap');
html,body,[class*="css"]{font-family:'EB Garamond','Cairo',Georgia,serif;background-color:#F7F3EE;}
.stApp{background-color:#F7F3EE;}
.q-card{background:white;border:1px solid #DDD5C8;border-radius:3px;padding:14px 18px 8px;
        margin-bottom:10px;box-shadow:0 1px 3px rgba(28,25,23,0.06);}
.q-num{font-size:10px;font-weight:600;color:#8B7355;letter-spacing:.1em;text-transform:uppercase;margin-bottom:4px;}
.q-text{font-size:15px;color:#1C1917;line-height:1.5;margin-bottom:10px;}
div[data-testid="stRadio"]>label{display:none;}
div[data-testid="stRadio"]>div{gap:8px!important;flex-direction:row!important;flex-wrap:wrap!important;}
div[data-testid="stRadio"]>div>label{background:#F7F3EE!important;border:1.5px solid #DDD5C8!important;
    border-radius:2px!important;padding:6px 16px!important;font-size:12px!important;color:#8B7355!important;
    font-family:'EB Garamond',Georgia,serif!important;cursor:pointer!important;transition:all 0.12s!important;}
div[data-testid="stRadio"]>div>label:has(input:checked){background:#2D2926!important;color:white!important;border-color:#2D2926!important;}
div[data-testid="stTextInput"] input{background:white!important;border:1px solid #DDD5C8!important;
    border-radius:2px!important;font-family:'EB Garamond','Cairo',Georgia,serif!important;}
.stButton>button{background:#2D2926!important;color:white!important;border:none!important;
    border-radius:2px!important;font-size:.95rem!important;font-weight:500!important;
    font-family:'EB Garamond',Georgia,serif!important;letter-spacing:.04em!important;
    box-shadow:none!important;transition:all 0.15s!important;}
.stButton>button:hover{background:#8B7355!important;}
.progress-wrap{background:#DDD5C8;height:2px;border-radius:2px;margin:6px 0 16px 0;}
.progress-fill{background:#8B7355;height:2px;border-radius:2px;transition:width 0.3s;}
</style>""", unsafe_allow_html=True)

# Session state
for k,v in [("lang","en"),("responses",{}),("report_done",False),
            ("access_granted",False),("access_error","")]:
    if k not in st.session_state: st.session_state[k]=v

# ── HEADER (shown always) ──
def show_header():
    cl,ch=st.columns([1,5])
    with cl:
        if os.path.exists(LOGO_FILE): st.image(LOGO_FILE,width=110)
    with ch:
        st.markdown("""<div style="padding:8px 0 0 0;">
            <div style="font-size:1.55rem;font-weight:500;color:#1C1917;
                        font-family:'EB Garamond',Georgia,serif;letter-spacing:.01em;line-height:1.2;">
                Child Behavior Checklist — CBCL/6-18</div>
            <div style="font-size:.8rem;color:#8B7355;margin-top:3px;letter-spacing:.08em;text-transform:uppercase;">
                Achenbach System of Empirically Based Assessment · Ages 6–18</div>
        </div>""", unsafe_allow_html=True)
    st.markdown('<hr style="border:none;border-top:1px solid #DDD5C8;margin:12px 0 16px;">',unsafe_allow_html=True)

show_header()

# ══════════════════════════════════════════════════════════════
#  ACCESS GATE
# ══════════════════════════════════════════════════════════════
if not st.session_state.access_granted:
    _,gc,_=st.columns([1,2,1])
    with gc:
        st.markdown("""
        <div style="background:white;border:1px solid #DDD5C8;border-radius:3px;
                    padding:2rem 2.4rem;text-align:center;">
            <div style="font-size:1.2rem;font-weight:500;color:#1C1917;
                        font-family:'EB Garamond',Georgia,serif;margin-bottom:.4rem;">Access Required</div>
            <div style="font-size:.85rem;color:#8B7355;margin-bottom:1.6rem;line-height:1.7;">
                Please enter the access code provided by your clinician to begin.<br>
                <span style="font-family:'Cairo',sans-serif;">
                أدخل رمز الوصول الذي زوّدك به الأخصائي للبدء.</span></div>
        </div>""", unsafe_allow_html=True)
        code_inp=st.text_input("Access code",placeholder="Enter access code",
                               type="password",key="code_inp",label_visibility="collapsed")
        if st.session_state.access_error:
            st.markdown(f'<div style="color:#C62828;font-size:.85rem;text-align:center;'
                        f'margin-top:.5rem;">{st.session_state.access_error}</div>',unsafe_allow_html=True)
        if st.button("Enter / دخول",use_container_width=True):
            raw_codes=st.secrets.get("ACCESS_CODE","")
            valid=[c.strip() for c in raw_codes.split(",") if c.strip()]
            if code_inp.strip() in valid:
                st.session_state.access_granted=True; st.session_state.access_error=""; st.rerun()
            else:
                st.session_state.access_error="Invalid access code. Please try again. / رمز الوصول غير صحيح."; st.rerun()
    st.stop()

# ══════════════════════════════════════════════════════════════
#  DONE SCREEN
# ══════════════════════════════════════════════════════════════
if st.session_state.report_done:
    _,cc,_=st.columns([1,3,1])
    with cc:
        st.markdown("""
        <div style="background:white;border:1px solid #DDD5C8;border-radius:3px;
                    padding:3rem 2rem;text-align:center;margin-top:2rem;">
            <div style="font-size:2.5rem;margin-bottom:1rem;">✓</div>
            <div style="font-size:1.4rem;font-weight:500;color:#1C1917;
                        font-family:'EB Garamond',Georgia,serif;margin-bottom:.8rem;">
                Thank you. Assessment submitted successfully.</div>
            <div style="font-size:1rem;color:#8B7355;line-height:1.9;margin-bottom:.6rem;">
                Your responses have been recorded and the report<br>has been sent to the clinician.</div>
            <div style="font-family:'Cairo',sans-serif;font-size:.95rem;color:#8B7355;line-height:1.9;">
                شكراً لك. تم تسجيل إجاباتك وإرسال التقرير إلى الأخصائي.</div>
        </div>""", unsafe_allow_html=True)
    st.stop()

# ══════════════════════════════════════════════════════════════
#  LANGUAGE TOGGLE — text only, no flags
# ══════════════════════════════════════════════════════════════
lang=st.session_state.lang
c1,c2,c3=st.columns([2,2,8])
with c1:
    if st.button("English",use_container_width=True):
        st.session_state.lang="en"; st.session_state.responses={}
        st.session_state.report_done=False; st.rerun()
with c2:
    if st.button("العربية",use_container_width=True):
        st.session_state.lang="ar"; st.session_state.responses={}
        st.session_state.report_done=False; st.rerun()
st.markdown(f'<div style="font-size:.72rem;color:#8B7355;letter-spacing:.07em;'
            f'text-transform:uppercase;margin-bottom:18px;">'
            f'{"English mode active" if lang=="en" else "النسخة العربية نشطة"}'
            f'</div>',unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  INTRO
# ══════════════════════════════════════════════════════════════
if lang=="en":
    st.markdown("""
    <div style="background:white;border:1px solid #DDD5C8;border-radius:3px;
                padding:1.4rem 1.8rem;margin-bottom:1.5rem;line-height:1.9;">
        <div style="font-size:.72rem;font-weight:600;color:#8B7355;letter-spacing:.1em;
                    text-transform:uppercase;margin-bottom:.6rem;">About This Assessment</div>
        <p style="font-size:.95rem;color:#1C1917;margin:0 0 .7rem 0;">
            The <strong>Child Behavior Checklist (CBCL/6-18)</strong> is a parent-completed questionnaire
            that assesses emotional and behavioral patterns in children aged 6 to 18.
            It is one of the most widely researched and validated tools in child and adolescent mental health.</p>
        <p style="font-size:.95rem;color:#1C1917;margin:0 0 .7rem 0;">
            <strong>How to answer:</strong> Think about your child's behavior
            <strong>now or during the past 6 months</strong> and choose:</p>
        <div style="font-size:.92rem;color:#1C1917;padding-left:.5rem;line-height:2;">
            <strong>0</strong> — Not True (as far as you know)<br>
            <strong>1</strong> — Somewhat or Sometimes True<br>
            <strong>2</strong> — Very True or Often True
        </div>
        <p style="font-size:.85rem;color:#8B7355;margin:.8rem 0 0 0;font-style:italic;">
            Please answer every item honestly, even if some do not seem to apply.
            There are no right or wrong answers. All responses are confidential.</p>
    </div>""", unsafe_allow_html=True)
else:
    st.markdown("""
    <div style="background:white;border:1px solid #DDD5C8;border-radius:3px;
                padding:1.4rem 1.8rem;margin-bottom:1.5rem;line-height:1.9;direction:rtl;text-align:right;">
        <div style="font-size:.72rem;font-weight:600;color:#8B7355;letter-spacing:.1em;
                    text-transform:uppercase;margin-bottom:.6rem;">نبذة عن هذا التقييم</div>
        <p style="font-size:.95rem;color:#1C1917;margin:0 0 .7rem 0;">
            <strong>قائمة أكنباك للسلوك الطفلي (CBCL/6-18)</strong> استبيان يُكمله ولي الأمر لتقييم الأنماط
            الانفعالية والسلوكية لدى الأطفال من سن ٦ إلى ١٨ سنة. وتُعدّ من أكثر الأدوات
            استخداماً وموثوقيةً في مجال الصحة النفسية للأطفال حول العالم.</p>
        <p style="font-size:.95rem;color:#1C1917;margin:0 0 .7rem 0;">
            <strong>طريقة الإجابة:</strong> تذكّر سلوك طفلك
            <strong>الآن أو خلال الستة أشهر الماضية</strong> واختر الأنسب:</p>
        <div style="font-size:.92rem;color:#1C1917;padding-right:.5rem;line-height:2;">
            <strong>0</strong> — غير صحيح (على حد علمك)<br>
            <strong>1</strong> — صحيح أحياناً<br>
            <strong>2</strong> — صحيح غالباً أو دائماً
        </div>
        <p style="font-size:.85rem;color:#8B7355;margin:.8rem 0 0 0;font-style:italic;">
            أجب بصدق على كل عبارة حتى تلك التي لا تبدو وثيقة الصلة بطفلك.
            لا توجد إجابات صحيحة أو خاطئة، وجميع الإجابات سرية تماماً.</p>
    </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  INTAKE FORM
# ══════════════════════════════════════════════════════════════
if lang=="en":
    st.markdown('<div style="font-size:.72rem;font-weight:600;color:#8B7355;letter-spacing:.08em;'
                'text-transform:uppercase;margin-bottom:.8rem;padding-bottom:.4rem;'
                'border-bottom:1px solid #DDD5C8;">Child Information</div>',unsafe_allow_html=True)
    c1,c2,c3=st.columns(3)
    with c1:
        child_name=st.text_input("Child's Full Name",placeholder="First and Last Name",key="cn")
        child_age =st.text_input("Age",placeholder="e.g. 10",key="ca")
    with c2:
        child_gender=st.radio("Gender",["Male","Female"],key="cg",horizontal=True)
        child_grade =st.text_input("School Grade",placeholder="e.g. Grade 4",key="cgr")
    with c3:
        rater       =st.text_input("Rater's Full Name",placeholder="Name",key="rn")
        relationship=st.text_input("Relationship to Child",placeholder="e.g. Mother, Father, Grandparent",key="rel")
    SCALE_OPTS=["0 — Not True","1 — Somewhat True","2 — Very True"]
    ITEMS=ITEMS_EN; item_label="Item"; dir_s="ltr"; align_s="left"
else:
    st.markdown('<div style="font-size:.72rem;font-weight:600;color:#8B7355;letter-spacing:.08em;'
                'text-transform:uppercase;margin-bottom:.4rem;padding-bottom:.4rem;'
                'border-bottom:1px solid #DDD5C8;direction:rtl;text-align:right;">'
                'بيانات الطفل</div>',unsafe_allow_html=True)
    st.caption("يرجى ملء جميع الحقول باللغة الإنجليزية  /  Please fill all fields in English")
    c1,c2,c3=st.columns(3)
    with c1:
        child_name=st.text_input("Child's Full Name / اسم الطفل",placeholder="First and Last Name",key="cn")
        child_age =st.text_input("Age / السن",placeholder="e.g. 10",key="ca")
    with c2:
        child_gender=st.radio("Gender / النوع",["Male","Female"],key="cg",horizontal=True)
        child_grade =st.text_input("School Grade / الصف",placeholder="e.g. Grade 4",key="cgr")
    with c3:
        rater       =st.text_input("Rater's Full Name / اسم المُقيِّم",placeholder="Name",key="rn")
        relationship=st.text_input("Relationship to Child / صلة القرابة",
                                   placeholder="e.g. Mother, Father, Grandparent",key="rel")
    SCALE_OPTS=["0 — غير صحيح","1 — أحياناً","2 — غالباً"]
    ITEMS=ITEMS_AR; item_label="بند"; dir_s="rtl"; align_s="right"

st.markdown("<br>",unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
#  ALL 120 RESPONSE SLOTS
# ══════════════════════════════════════════════════════════════
responses=st.session_state.responses
all_answered=True

for key in ITEM_KEYS:
    item_text=ITEMS[key]
    display_num=str(key)

    st.markdown(f'<div class="q-card" style="direction:{dir_s};">'
                f'<div class="q-num" style="text-align:{align_s};">{item_label} {display_num}</div>'
                f'<div class="q-text" style="text-align:{align_s};">{item_text}</div>'
                f'</div>',unsafe_allow_html=True)

    saved=responses.get(key)
    choice=st.radio(f"item_{key}",SCALE_OPTS,index=saved,key=f"resp_{key}",
                    horizontal=True,label_visibility="collapsed")
    if choice is None: all_answered=False
    else:
        val=int(choice[0]); responses[key]=val; st.session_state.responses[key]=val

# ══════════════════════════════════════════════════════════════
#  PROGRESS & SUBMIT
# ══════════════════════════════════════════════════════════════
total_items=len(ITEM_KEYS)
answered_count=len([v for v in responses.values() if v is not None])
pct=int((answered_count/total_items)*100)
prog_text=(f"{answered_count} of {total_items} answered" if lang=="en"
           else f"{answered_count} من {total_items} بنداً")

st.markdown(f'<div style="text-align:center;font-size:.78rem;color:#8B7355;'
            f'letter-spacing:.06em;margin-top:1.5rem;">{prog_text}</div>'
            f'<div class="progress-wrap"><div class="progress-fill" style="width:{pct}%"></div></div>',
            unsafe_allow_html=True)

if not all_answered and answered_count>0:
    warn=("⚠ Please answer all items before submitting." if lang=="en"
          else "⚠ يرجى الإجابة على جميع البنود قبل الإرسال.")
    st.markdown(f'<div style="background:#FFF8F0;border-left:3px solid #E07B39;'
                f'padding:1rem 1.2rem;border-radius:0 2px 2px 0;'
                f'font-size:.88rem;color:#7A3D1A;margin:1rem 0;">{warn}</div>',
                unsafe_allow_html=True)

st.markdown("<br>",unsafe_allow_html=True)
btn_label="Submit" if lang=="en" else "إرسال"
cb,_=st.columns([2,3])
with cb:
    submit=st.button(btn_label,use_container_width=True,disabled=(answered_count<total_items))

if submit and answered_count>=total_items:
    child_name_v = child_name  or "Child"
    child_age_v  = child_age   or "—"
    gender_v     = child_gender
    rater_v      = rater       or "Parent/Caregiver"
    rel_v        = relationship or ""
    responses_v  = dict(st.session_state.responses)

    with st.spinner("⏳ Processing…" if lang=="en" else "⏳ جاري المعالجة…"):
        scores    = compute_scores(responses_v)
        report_en = generate_report_en(child_name_v,child_age_v,gender_v,rater_v,rel_v,scores,responses_v)
        charts    = {"syndrome":make_syndrome_chart(scores),"broadband":make_broadband_chart(scores),
                     "dsm":make_dsm_chart(scores),"pie":make_pie_chart(responses_v)}
        fn_pdf    = f"{child_name_v.replace(' ','_')}_CBCL_EN.pdf"
        buf_pdf   = build_pdf(report_en,scores,charts,child_name_v,child_age_v,gender_v,rater_v,rel_v,responses_v)
        try:
            send_email(child_name_v,buf_pdf,fn_pdf,scores)
        except Exception:
            pass

    st.session_state.report_done=True
    st.rerun()
