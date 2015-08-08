#!/bin/bash
if [ $# -ne 5 ]; then
   echo "Usages: fetchall <branch> <begintag> <endtag> <verno>"
   echo "Example: fetchall 1505N GTITAG1 GITTAG2 2.52.01.01 ctie"
   exit 1
fi
branch=$1
begintag=$2
endtag=$3
verno=$4
moudle=$5

#config
tagpath=/data/luowb/ctie/tag
LOG_FILE=/data/luowb/ctie/log.txt

codesum=$tagpath/TAG_${verno}_${begintag}_${endtag}.csv

#log output
log() {
   log_content="`date +'%F %T'` $*"
   echo "${log_content}"
   echo "${log_content}" >> ${LOG_FILE}
}

warn() {
   log "[WARNING] $*"
   exit
}
cd $branch/$moudle

echo "" > $codesum

log "git fetch --pregress --all"
git fetch --progress --all

log "parse git diff"
revlist=$(git log  $begintag..$endtag --pretty=format:%H --perl-regexp --author='^(?!(.*(kfzx-sujh|KFZX-ZHANGY11|kfzx-wuyl|kfzx-zhangyang)))' --diff-filter=MCRA)


formatgit(){
	content=$*
 	echo -n "${content}" >> $codesum
}

for rev in $revlist
do
	files=$(git log -1 --pretty=format: --name-only $rev)
	formatgit "@@@@"
	formatgit "$(git log -1 --pretty=%H ${rev})"
	formatgit "@@@@,@@@@"
        formatgit "$(git log -1 --pretty=%s ${rev})"
        formatgit "@@@@,@@@@"
        formatgit "$(git log -1 --pretty=%an ${rev})"
        formatgit "@@@@,@@@@"
        formatgit "$(git log -1 --pretty=%b ${rev})"
        formatgit "@@@@,@@@@"
        formatgit "F-CTIE V${verno}" 
        formatgit "@@@@,@@@@"
	for file in $files
	do
		echo "$file" | grep -q -E "Cmp_CTIE-UTIL3|189_server|189_client|Cmp_CTIE-COMPILE3|com.icbc.ctie.environment|commit-msg|reviewer|.vcxproj"
		if [ $? -ne 0 ];then
			formatgit "$file"
			echo -e "\n" >> $codesum
		fi
	done
	formatgit "@@@@"
	echo -e ",\n" >> $codesum
done

#git log $begintag..$endtag --pretty=format:'@@@@%H@@@@%s@@@@,@@@@%an@@@@,@@@@%b@@@@,' --perl-regexp --author='^(?!(.*(kfzx-sujh|kfzx-luowb|KFZX-ZHANGY11|kfzx-wuyl|kfzx-zhangyang)))' > $codesum

log "$codesum is created"
sed -i s/\"/\"\"/g $codesum
sed -i s/@@@@/\"/g $codesum
sed -i /^$/d  $codesum

sql="db.ctievers.insert({name:\"$branch\",verno:\"F-CTIE V$verno\",stag:\"$begintag\",etag:\"$endtag\"})";
log "execute command: $sql "
echo "$sql" | /data/luowb/mongo/bin/mongo 127.0.0.1:3001/meteor --shell
log "Format ${codesum} finished"


if [ -e ${codesum} ]; then
/data/luowb/mongo/bin/mongoimport  -d meteor -c tasks --type csv -f "hash,title,coder,detail,verno,codelist" --file ${codesum} --host localhost --port 3001
fi

log "Import test list to  MongoDB finished"
