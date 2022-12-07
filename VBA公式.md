ファイル名前を取り出す
=MID(CELL("filename"),FIND("_FY",CELL("filename"))+3,FIND(".xlsm",CELL("filename"))-FIND("_FY",CELL("filename"))-3)
