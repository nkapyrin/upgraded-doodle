#encoding:utf-8

from lxml import etree # pip install lxml
import sympy as sp
from sympy.printing.latex import latex
import codecs

db = etree.parse("db.xml").getroot();
template = u""
with codecs.open( 'tex/template.tex', 'r', "utf-8") as f:
    template = f.read()

# Информация о задании
basic_info = { u"#INSTRUMENT_NAME#":     db.find('instrument').attrib["name"], \
               u"#INSTRUMENT_NAME_ACC#": db.find('instrument').attrib["name_acc"], \
               u"#EFFECT_NAME#":         db.find('effect').attrib["name"], \
               u"#EFFECT_NAME_ACC#":     db.find('effect').attrib["name_acc"], \
               u"#LIMITATIONS#":         db.find('limitation').attrib["name"], \
               u"#LIMITATIONS_ACC#":     db.find('limitation').attrib["name_acc"], \
               u"#YOUR_STATUS#":         db.find('author').attrib["status"], \
               u"#YOUR_GROUP#":          db.find('author').attrib["group"], \
               u"#YOUR_NAME#":           db.find('author').attrib["name"] \
             }


# Информация из общей базы данных, зависящая от выбора задания
general_info = { u"#INSTRUMENT_DESCRIPTION#": db.find("description[@name='%s']" % basic_info['#INSTRUMENT_NAME#']).attrib['text'], \
                 u"#EFFECT DESCRIPTION#":     db.find("description[@name='%s']" % basic_info['#EFFECT_NAME#']).attrib['text'], \
                 u"#EFFECT_VARIABLES_AND_DESCRIPTIONS#":     db.find("model[@name='%s']" % basic_info['#EFFECT_NAME#']).attrib['description'], \
                 u"#INSTRUMENT_VARIABLES_AND_DESCRIPTIONS#": db.find("model[@name='%s']" % basic_info['#INSTRUMENT_NAME#']).attrib['description'] }


math_models = { u"#INSTRUMENT_MODEL#": [s.strip() for s in db.find("model[@name='%s']" % basic_info['#INSTRUMENT_NAME#']).attrib['math'].split('=')],\
                u"#EFFECT_MODEL#":     [s.strip() for s in db.find("model[@name='%s']" % basic_info['#EFFECT_NAME#']).attrib['math'].split('=')] }


# Получение модели погрешности
outVar =    sp.Symbol( math_models['#INSTRUMENT_MODEL#'][0] )
deviceEqn = sp.sympify( math_models['#INSTRUMENT_MODEL#'][1] )
effectVar = sp.Symbol( math_models['#EFFECT_MODEL#'][0] )
diffVars =  list( (sp.sympify( math_models['#EFFECT_MODEL#'][1] )).free_symbols )
effectEqn = sp.sympify( math_models['#EFFECT_MODEL#'][1] )
delta_outVar = sp.Symbol( 'Delta_' + math_models['#INSTRUMENT_MODEL#'][0] )
delta_diffVars = [ sp.Symbol( 'Delta_' + str(d) ) for d in diffVars ]
subsEqn = deviceEqn.subs( effectVar, effectEqn )
diffEqn = sp.sympify( '0' )
for s,ds in zip( diffVars,delta_diffVars):
    diffEqn = diffEqn + sp.diff(subsEqn, s) * ds
math_models['#INSTRUMENT_ERROR_MATH#'] = [ delta_outVar, diffEqn ]


# Перевод уравнений в формат LaTeX
math_latex = {}
for k in math_models.keys():
    LHS_sp,  RHS_sp  = sp.sympify( math_models[k][0] ), sp.sympify( math_models[k][1] )
    math_latex[k] = sp.latex(LHS_sp) + ' = ' + sp.latex(RHS_sp)


# Заменить все ключевые выражения
for d in [ basic_info, general_info ]:
    for k in d.keys(): template = template.replace( k, d[k].replace('_','\_').replace(u'±',u'$\pm$').replace(u'—',u'--') )

for d in [ math_latex ]:
    for k in d.keys(): template = template.replace( k, d[k] )

# Добавить БД задания в виде QR кодов
# 1249 -- количество символов которое помещается в QR Version 20 при уровне коррекции ошибок L
import qrcode
xml_text = etree.tostring( db, encoding='UTF-8' ).decode("utf-8")
qr_txt = u""
for i,s in enumerate([ xml_text[i:i+1249] for i in range(0, len(xml_text), 1249) ]):
    qr = qrcode.QRCode( version=1, box_size=2, border=4, error_correction = qrcode.constants.ERROR_CORRECT_L )
    qr.add_data( s )
    qr.make( fit=True )
    im = qr.make_image()
    im.save( "tex/qr_tmp/db_qr_%d.png" % i )
    qr_txt = qr_txt + u"\includegraphics[width=.4\\textwidth]{qr_tmp/db_qr_%d.png}\\\\" % i
template = template.replace( u"#QR_DUMP_HERE#", qr_txt )

#for i,s in enumerate([ xml_text[i:i+1249] for i in range(0, len(xml_text), 1249) ]):
#    qr_txt = qr_txt + u"\qrcode{%s}" % s

# Записать выходной файл
filename = "tex/result.tex"
print filename
with codecs.open( filename, "w", "utf-8") as f:
    f.write( template )

import os
os.chdir("tex")
os.system( 'pdflatex -interaction=batchmode result.tex' ) # -synctex=1

