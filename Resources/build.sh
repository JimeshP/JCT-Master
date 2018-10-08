mv admin.war admin_bak.war
mv user.war user_bak.war
cp /Users/rajan/Documents/dev/workspace/jct/JCT/target/user.war .
cp /Users/rajan/Documents/dev/workspace/jct/JCTAdmin/target/admin.war .
jar cvf app_7.zip admin.jar user.jar ROOT.jar
