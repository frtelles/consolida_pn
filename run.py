from helpers import *
import constants as c

if __name__ == "__main__":

    # check if data directory exists
    check_directory(c.ROOT_DIR, c.DIRECTORY)

    pn_consolidado_header(c.NAME_FILE)

    tipo_pn, line_id, data_inc, regional, hostname, site, swap_de, projeto,\
        sub_projeto, anel_ipran, anel_ipnodeb, integrado, fabricante, integrador, aplicacao,\
        model_bastidor, model_simp, municipio, remanejamento, comentario,\
        projeto_capex, sub_projeto_capex, si_capex, ano_imp = le_pn_ipran()

    pn_consolidado_escreve_ipran(c.NAME_FILE, tipo_pn, line_id, data_inc, regional, hostname,
                                 site, swap_de, projeto, sub_projeto, anel_ipran, anel_ipnodeb,
                                 integrado, fabricante, integrador, aplicacao, model_bastidor, model_simp,
                                 municipio, remanejamento, comentario, projeto_capex, sub_projeto_capex, si_capex, ano_imp)

    tipo_pn, line_id, data_inc, regional, hostname, site, projeto, sub_projeto, \
        anel_ipran, anel_ipnodeb, integrado, fabricante, integrador, aplicacao, model_bastidor,\
        comentario, projeto_capex, sub_projeto_capex, si_capex, ano_imp, executado = le_pn_upgrade()

    pn_consolidado_escreve_upgrade(c.NAME_FILE, tipo_pn, line_id, data_inc, regional, hostname, site, projeto,
                                   sub_projeto, anel_ipran, anel_ipnodeb, integrado, fabricante, integrador, aplicacao,
                                   model_bastidor, comentario, projeto_capex, sub_projeto_capex, si_capex, ano_imp, executado)
    print('Done!')
