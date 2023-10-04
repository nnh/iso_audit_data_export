#!/bin/env bash
sjis="charset=unknown-8bit"
filename[0]="ClassEditPath.cls"
filename[1]="ClassFolderPathManager.cls"
filename[2]="ConstantsModule.bas"
filename[3]="ConvertToPdf.bas"
filename[4]="CreateText.bas"
filename[5]="ExportVba.bas"
filename[6]="FileUtils.bas"
filename[7]="Utils.bas"


filepath=".././programs/vba/modules/"
for fname in ${filename[@]}; do
	echo ${filepath}${fname}
	temp=$(file -i ${filepath}${fname} |awk '{print $3}')
	echo ${temp}
	if [ $temp = $sjis ]; then
  		iconv -f SHIFT-JIS -t UTF-8 ${filepath}${fname} > ${filepath}temp.bas
  		echo "SHIFT-JIS -> UTF-8"
	else
  		iconv -f UTF-8 -t SHIFT-JIS ${filepath}${fname} > ${filepath}temp.bas
  		echo "UTF-8 -> SHIFT-JIS"
	fi
	cp ${filepath}temp.bas ${filepath}${fname}
	rm ${filepath}temp.bas
done

