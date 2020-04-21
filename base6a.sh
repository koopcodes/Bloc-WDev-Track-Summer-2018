# a million commits
for Y in {2017..2017}
do
  mkdir $Y
  cd $Y
  for M in {01..12}
  do
    mkdir $M
    cd $M
    for D in {06..06}
    do
      mkdir $D
      cd $D
      for i in {01..04}
      do
        echo "/$i on $M/$D/$Y" > glob.js
		export GIT_AUTHOR_NAME="koopcodes"
		export GIT_AUTHOR_EMAIL="koop@koop.codes"
		export GIT_AUTHOR_DATE="$Y-$M-$D 12:$i:00"
		export GIT_COMMITTER_NAME="koopcodes"
		export GIT_COMMITTER_EMAIL="koop@koop.codes"
        export GIT_COMMITTER_DATE="$Y-$M-$D 12:$i:00"
        git add glob.js -f
        git commit --date="$Y-$M-$D 12:0$i:00" -m "$i on $M $D $Y"
      done
      cd ../
    done
    cd ../
  done
  cd ../
done
git push origin master
git rm -rf 20**
# git rm -rf 19**
git commit -am "cleanup"
git push origin master
