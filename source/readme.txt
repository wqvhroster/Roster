Checkout code:
git clone https://github.com/jeremyrixon/roster.git

cd roster

Install ca cert:
open certs/ca.crt -> Add. Choose Keychain "login" and category "Certificates". Find "localhost-ca" and double-click. Trust -> when using this certificate -> always trust. Close. Enter password.

Install node

Setup npm:
sudo npm install -g yo generator-office webpack webpack-dev-server webpack-cli
npm install

Setup VSCode:
Add folder to workspace -> git/roster

Dist:
webpack

Open Roster:
https://wqvet-my.sharepoint.com/:x:/r/personal/shan_westqueanbeyanvet_com_au/_layouts/15/Doc.aspx?sourcedoc=%7B273464bc-4774-4c69-aa58-be2242637e50%7D&action=default&uid=%7B273464BC-4774-4C69-AA58-BE2242637E50%7D&ListItemId=46&ListId=%7B8367E97F-C77E-40E2-A759-CA6C337C843D%7D&odsp=1&env=prod

Side-load app:
Insert -> Office add-ins -> Upload My add-in -> Browse -> (find the manifest.xml) -> upload -> show task pane
