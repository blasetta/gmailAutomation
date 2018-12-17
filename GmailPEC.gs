/* Parameters to be customized
   The code is written in a simple way
 */

var mailFatturefrom="xy@fatture.it";
var myPEC="myPEC@legalmail.it";
var myGmailAccount="myGMail@gmail.com";
var PECSDI="sdi@pec.it";
var mailnotifica="maymail@mycompany.com";


/* Functions List
 TrovaMails - finds, labels nd marks invoices
 SendtoSDI  - send the email with the invoice
 Report     - Activity Report for all the Invoices in the last month
 gScarti    - Refused email handling

 */


/* Report Send Receive "Fatture Elettroniche"  */
function Report() {
    /* Variables   */
    var gmese=31; var unmesefa=Math.floor(new Date(new Date()-gmese*24*60*60*1000).getTime()/1000);
    var searchstring= 'is:sent has:attachment to:'+PECSDI+' newer:' +unmesefa;
    var JReport={};
    var risp='';

    // Look for all the invoices
    var msgs1=TrovaMails(searchstring);
    if (msgs1.length==0) {risp="No Data Found"; return risp;}

    for (var i1 = 0 ; i1 < msgs1.length; i1++) {
        var currmsg=msgs1[i1];
        // attachment nane
        var nomefile=currmsg.getAttachments()[0].getName()||'noname';
        if (!JReport[nomefile]) JReport[nomefile]={};
        var JR = JReport[nomefile];
        JR.subject=currmsg.getSubject(); JR.sent=currmsg.getDate(); JR.sent=[JR.sent.getDate(), JR.sent.getMonth()+1, JR.sent.getFullYear()].join('/');
    }

    // SDI Notifications
    searchstring= 'from:'+PECSDI+' subject:(notifica || mancata)  newer:' +unmesefa;
    var msgs2=TrovaMails(searchstring);

    for (var i2 = 0 ; i2 < msgs2.length; i2++) {
        /* File name-  xml fatture  */
        var currmsg=msgs2[i2];
        var nomefile=(currmsg.getBody().split('.xml'))[0]; nomefile='IT'+nomefile.split('Il file IT')[1]+'.xml';
        if (!JReport[nomefile]) JReport[nomefile]={};
        var JR = JReport[nomefile]; var Oggetto=currmsg.getSubject();
        JR.dataesito=currmsg.getDate(); JR.dataesito=[JR.dataesito.getDate(), JR.dataesito.getMonth()+1, JR.dataesito.getFullYear()].join('/');

        // outcomes
        if (Oggetto.indexOf("scarto")>0) { JR.esito="scarto";}
        else if (Oggetto.indexOf("esito")>0) {JR.esito="OK";}
        else if (Oggetto.indexOf("mancata")>0) {JR.esito="mancata consegna";}

    }
    /* HTML Table */
    var hBody='';
    for(var fn in JReport) {
        var JR=JReport[fn];
        hBody+='<tr><td><b>'+fn+'</b></td><td>'+JR.subject+'</td><td>'+JR.sent||''+'</td><td>'+JR.esito||''+'</td><td>'+JR.dataesito||''+'</td></tr>'
    }
    hBody='<h3>Report Fatturazione</h3><table><tr><th>Nome File</th><th>Oggetto</th>Data<th></th><th>Esito</th><th>Dataesito</th></tr>'+hBody+'</table>';

    var blob = Utilities.newBlob(hBody, 'text/html', 'Report.xls');

    /* Send the notification to the business mail (mailnotifica)   */
    GmailApp.sendEmail(mailnotifica, 'Report fatturazione elettronica del mese', "..",  {
        htmlBody: hBody,
        attachments: [blob]
    });

    return hBody;

}

/* All the new Invoices ************/
function SendtoSDI() {
    var cerca='in:inbox has:attachment -label:inviate from:'+mailFatturefrom;   var conta=0;

    var msgs=TrovaMails(cerca, [ "inviate", "processed"],null, ["archive"]);
    Logger.log(' messaggi trovati   '+msgs.length);

    for (var i = 0 ; i < msgs.length; i++) {
        var currmsg=msgs[i];   var work=currmsg.getSubject().split(' PEC:');  var Oggetto=work[0]; mailfrom=work[1];
        Logger.log(' Oggetto  '+currmsg.getSubject());
        var attachments = currmsg.getAttachments();
        GmailApp.sendEmail(PECSDI, Oggetto, "..",  {
            from: myPEC,
            htmlBody: currmsg.getBody(),
            attachments: attachments
        });
        conta+=1;

    }
    var msgsElab=TrovaMails('is:sent has:attachment to:'+PECSDI+' -label:inAttesa', ["inAttesa"], null, ["star"]);
    return 'Elaborate '+conta+' mail';
}


/* Management of refused Invoices  */
function gScarti() {
    var cerca='in:inbox subject:(POSTA CERTIFICATA: Notifica di scarto) -label:processed '; var resp='';
    var msgs=TrovaMails(cerca, ["processed"],null, ["archive"]);
    Logger.log(' messaggi trovati   '+msgs.length); resp+=' Scarti trovati   '+msgs.length;

    for (var i = 0 ; i < msgs.length; i++) {
        var currmsg=msgs[i]; var msgbody=currmsg.getBody(), nomefile=''; var s=msgbody.split('.xml'); var trovato=false; var documenti=null;
        var newmailSubject=currmsg.getSubject();
        var newmailBody=currmsg.getBody();
        var newmailAllegati=currmsg.getAttachments();

        nomefile='IT'+s[0].split('Il file IT')[1]+'.xml';
        /* Trova il numero del documento e ritornalo  */
        if (!nomefile) { Logger.log(' non ho trovato il file XML! '+msgbody); newmailSubject=newmailSubject+' NON HO TROVATO il NOME DEL FILE!  '; }
        else {
            var cercaInvio='is:sent has:attachment to:'+PECSDI+' '+nomefile;

            /* Find the mail with that document numeber */
            var msgs2=TrovaMails(cercaInvio,["scartate"],null, ["star"]);
            Logger.log(' Mail inviate con allegato Originali   '+msgs2.length); resp+=' Mail inviate con allegato Originali   '+msgs2.length;
            if (msgs2.length>0) {
                trovato=true;
                // prendi il documento allegato
                newmailSubject=newmailSubject+' da: '+msgs2[0].getSubject();
                newmailBody=newmailBody+'----------------------'+msgs2[0].getBody();
                var newAtt=msgs2[0].getAttachments();
                for (var aa = 0 ; aa < newAtt.length; aa++) { newmailAllegati.push(newAtt[aa]); }
            }
            else { newmailSubject=newmailSubject+' NON HO TROVATO il DOC ORIGINALE!  ';}

        } // end nome file
        /* Manda la notifica -- mail normale  */
        GmailApp.sendEmail(mailnotifica, newmailSubject, "..",  {
            htmlBody: newmailBody,
            attachments: newmailAllegati
        });

    }

    return resp;
}






/* Flag invoices as sent and accepted : labels */
function flagInvoice() {

    var cerca='in:inbox subject:(ACCETTAZIONE: OR CONSEGNA:) -label:processed ';

    var msgs=TrovaMails(cerca, ["processed"], null, ["archive","read"]); var resp='';
    Logger.log(' messaggi trovati   '+msgs.length);  resp+='Ricevute trovate '+msgs.length;

    for (var i = 0 ; i < msgs.length; i++) {
        /* Cerca la mail corrispondente */
        var newcerca='is:sent subject:'+msgs[i].getSubject().replace('ACCETTAZIONE: ','').replace('CONSEGNA: ','');
        var matchmsg=TrovaMails(newcerca, [ "accettate"],[ "inAttesa"], ["star"]); resp+=' - Messaggi corrispondenti  trovati '+matchmsg.length;
    }
    return resp;
}


/* TovaMails */

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

