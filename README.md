# **VBA SCRIPTS**
![vba](https://img.shields.io/badge/VBA-o365-green)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

# **Snip table to email script**
## **Steps to run**
1. Was only tested on the latest o365 version as of February 3rd, 2022.
> requires outlook library. In the VBA editor window (ALT+F11 on windows) click on tools and choose Microsoft Office 16.0.
<br>
</br>

![office](/assets/op.PNG)
1. This only works if you have Outlook email client installed on your local machine.
2. Code copies used range (not active range) to email.
```vba
'' Change to active range if you want to choose range manually
Set copyrng = ws.UsedRange
```
3. Raw files looks like this:
![input](/assets/od.PNG)
4. Output should look something like below:
![output](/assets/oo.PNG)
5. Copy code to a new VBA module and hit F5 to run code.

6. [Link](/scripts/tableToEmail.bas) to code.




