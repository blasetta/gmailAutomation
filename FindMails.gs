/* Snippet
   Look for Emails with the cerca (search) condition and set unset labels and actions 
   Returns an array of messages for further processing :)
*/

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
