# Miette is "small sweet thing" in french

from doc.reader import DocReader

r = DocReader('../tests/doc/mw_lorem_ipsum.doc')
#r = DocReader('../tests/doc/gd_lorem_ipsum.doc')
#r = DocReader('../tests/doc/oo_lorem_ipsum.doc')
#r = DocReader('../tests/doc/te_lorem_ipsum.doc')

#r = DocReader('../tests/doc/mw_vesna_yandex_ru.doc')
#r = DocReader('../tests/doc/gd_vesna_yandex_ru.doc')
#r = DocReader('../tests/doc/oo_vesna_yandex_ru.doc')
#r = DocReader('../tests/doc/te_vesna_yandex_ru.doc')

print r.read()
