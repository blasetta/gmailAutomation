# gmailAutomation
```Automate Gmail for Invoices via Secure Email```

With 1,2 Billion users (probably underestimated) Gmail is one of the most popular email services. It has a huge set of functions and lets you search tour emails like a datastore.
If you have to manage business information, some automation is a must. This contribution is a companion for a Medium article in Italian and addressed to Italy because we have electronic invoicing that may travel on secure email.
And Gmail may help for managing and automating all the related communications.

## Functions addressed

### Searching
If you search for something you just type in some words and Gmail will look for them in all your messages.
If you want to do something more complex, you open the dropdown menu like in the following figure and make the desired selections; 
all this selections may be supplied as text.   
![Search mask](https://cms-assets.tutsplus.com/uploads/users/988/posts/27445/image/Gmail-search(1).jpg)

[Laura Spencer explains everything ](https://business.tutsplus.com/tutorials/how-to-search-your-emails-in-gmail--cms-27445) and  [here the official site](https://support.google.com/mail/answer/7190?hl=en)

### Labelling
Labelling is the easiest way to classify. Label=TAG. The same stuff of Twitter. You decide the label and they magically appear on the menu on the left side. 

### What the hell is archiving?
When you manage automated email, their number may became overwhelming. If you don't to see them in your inbox, but without deleting them, the solution is archiving.

## Search and Classify with code
The procedure "FindMails" will let you find a precise set of messages and label or activate some action on them. The labels are supplied as text and created automatically if they don't already exist. 
Moreover you can mark them as important, or with stars or you may archive them.

```javascript

function TrovaMails(cerca, labels, nolabels, actions) {
    var listaMsg=[];
    var threads = GmailApp.search(cerca);
    if (!threads || threads.length==0) {Logger.log(' no data found '); return []; }

    /* Trovate mail! */
    var mylabels=[]; var mynolabels=[]; actions=actions||[];

    if (labels) for (var l1 = 0 ; l1 < labels.length; l1++) {
        mylabels.push(GmailApp.getUserLabelByName(labels[l1])||GmailApp.createLabel(labels[l1])) ;
    };
    if (nolabels) for (var nl = 0 ; nl < nolabels.length; nl++) {
        mynolabels.push(GmailApp.getUserLabelByName(nolabels[nl]) || GmailApp.createLabel(nolabels[nl])) ;
    };

    var msgs = GmailApp.getMessagesForThreads(threads);
    for (var i = 0 ; i < msgs.length; i++) { for (var j = 0; j < msgs[i].length; j++) {
        var msg=msgs[i][j]; listaMsg.push(msg);
        for (var x = 0 ; x < actions.length; x++) {
            if (actions[x]=="archive") msg.getThread().moveToArchive();
            else if (actions[x]=="read") msg.markRead();
            else if (actions[x]=="unread") msg.markUnread();
            else if (actions[x]=="star") msg.star();
            else if (actions[x]=="unstar") msg.unstar();
        }

    }}

    for (var t in threads) {
        for (var i = 0 ; i < mylabels.length; i++) threads[t].addLabel(mylabels[i]);
        for (var i = 0 ; i < mynolabels.length; i++) threads[t].removeLabel(mynolabels[i]);
    }

    return listaMsg;
}
```
## Search and Merge related email
When the document presents some formal errors, the Government refuse it with a message and a short explanation.

Sounds OK, but you don't have any other information about the original document BUT the file name (that have to be composed by your code + a counter).
> So, having the original message AND the invoice document would be nice, definitely.
Function gscarti does exactly that:

> Finds the refusals and return in msgs the messages
```
var cerca='in:inbox subject:(POSTA CERTIFICATA: Notifica di scarto) -label:processed '; var resp='';
    var msgs=TrovaMails(cerca, ["processed"],null, ["archive"]);
```    
> for each refusals finds the original message and stars it
```
 var cercaInvio='is:sent has:attachment to:'+PECSDI+' '+nomefile;
 var msgs2=TrovaMails(cercaInvio,["scartate"],null, ["star"]);
```  
> and merges main content and attchments
```
    newmailBody=newmailBody+'----------------------'+msgs2[0].getBody();
    var newAtt=msgs2[0].getAttachments();
     for (var aa = 0 ; aa < newAtt.length; aa++) { newmailAllegati.push(newAtt[aa]); }
```  

