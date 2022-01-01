/* global TrelloPowerUp */

var t = TrelloPowerUp.iframe();
var _debug = false;


// -- Want to know when you are being closed?
window.addEventListener('unload', function(e) {
  // Our board bar is being closed, clean up if we need to
});



// -- Action on export button

const button = document.getElementById('exportButton');
button.addEventListener('click', function(event){
    
  // Parse data  
  return Promise.all([t.board("all"), t.lists("all")]).then(function(result) {
    
    var board = result[0];
    var lists = result[1];
    var customFieldsRef = board.customFields;
    var customFieldsNameRef = new Array();
    var customFieldsNamePrefix = "CF | ";
    var cardsToExcludeFieldTarget = "Ne pas exporter";
    
    var num = 0;
    var trello_data = new Array();
    
    if(_debug) console.log(board);
    if(_debug) console.log(lists);
    
    lists.forEach((list) => {
      
      let posCard = 0;
      
      list.cards.forEach((card) => {
        
        if(_debug) console.log(card);
        
        // Card Labels
        let labels = new Array();
        card.labels.forEach((label) => {
          labels.push(label.name);
        });
        
        // Card Custom Fields
        let cFields = new Array();
        customFieldsRef.forEach((refCField) => {
          
          if( num == 0 && customFieldsNameRef.indexOf(customFieldsNamePrefix+refCField.name) === -1 ) {
            customFieldsNameRef.push(customFieldsNamePrefix+refCField.name);
          } 
          
          var cardCField = card.customFieldItems.find(f => f.idCustomField === refCField.id);
          var cardCFieldValue = "";
          
          // custome field value
          if(cardCField && cardCField.hasOwnProperty("idValue")) {
            // type combo value
            var option = refCField.options.find(o => o.id === cardCField.idValue);
            cardCFieldValue = ( option.value.text ? option.value.text : option.value.number );
          } else if(cardCField && cardCField.hasOwnProperty("value")) {
            if(cardCField.value.text) {
              // type text
              cardCFieldValue = cardCField.value.text;
            } else if(cardCField.value.number) {
              // type number
              cardCFieldValue = cardCField.value.number;
            } else if(cardCField.value.checked) {
              // type checkbox
              cardCFieldValue = "oui";
            } else {
              //unknown
            }
          }
          
          cFields[customFieldsNamePrefix+refCField.name] = cardCFieldValue;
          
        });
        
        // card to exclude if matching critera
        if( cFields[customFieldsNamePrefix+cardsToExcludeFieldTarget] == "oui" ) {
          if(_debug) console.log("card excluded : "+card.name);
          return;
        }
        // exclude column if matching critera
        if( cFields.indexOf(customFieldsNamePrefix+cardsToExcludeFieldTarget) > -1 ) {
          delete cFields[customFieldsNamePrefix+cardsToExcludeFieldTarget];
        }
        
        // Card Members
        let members = new Array();
        card.members.forEach((member) => {
          members.push(member.fullName);
        });
        
        trello_data.push(
            Object.assign(
              {
                "NUM": ++num,
                "LISTE":list.name,
                "POS": ++posCard,
                "CARTE TRELLO":card.name,
                "AFFECTATIONS":members.join(' \n'),
                "ETIQUETTES":labels.join(' \n'),
                "DESCRIPTION":card.desc,
                "URL":"https://trello.com/c/"+card.shortLink
              },cFields
            )
          );

      });
      
    });
    
    var wb = XLSX.utils.book_new();
    wb.Props = {
            Title: "Tableau Trello",
            Subject: "Tableau Trello",
            Author: "Trello",
            CreatedDate: new Date()
    };

    wb.SheetNames.push("Trello");

    if(_debug) console.log(trello_data);
    
    // exclude column if matching critera
    if( customFieldsNameRef.indexOf(customFieldsNamePrefix+cardsToExcludeFieldTarget) > -1 ) {
      customFieldsNameRef.splice(customFieldsNameRef.indexOf(customFieldsNamePrefix+cardsToExcludeFieldTarget), 1);
    }
    
    if(_debug) console.log(customFieldsNameRef);
    
    var ws = XLSX.utils.json_to_sheet(trello_data, {header:[...["NUM","LISTE","POS","CARTE TRELLO","AFFECTATIONS","ETIQUETTES"],...customFieldsNameRef,...["DESCRIPTION","URL"]]});
    ws['!cols'] = [{wch:4}, {wch:30}, {wch:4}, {wch:100}, {wch:30}, {wch:30}];
    customFieldsNameRef.forEach((col) => {
      ws['!cols'].push({wch:30});
    });
    ws['!cols'].push({wch:100}, {wch:30});
    
    wb.Sheets["Trello"] = ws;
    var wbout = XLSX.write(wb, {bookType:'xlsx',  type: 'binary'});

    function s2ab(s) {

            var buf = new ArrayBuffer(s.length);
            var view = new Uint8Array(buf);
            for (var i=0; i<s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
            return buf;

    }
    
    var today = new Date();
        
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), 'trello_'+today.getFullYear()+'-'+(today.getMonth()+1)+'-'+today.getDate()+'.xlsx');
    
  }); 
    
 
});



t.render(function(){

  // this function we be called once on initial load
  // and then called each time something changes that
  // you might want to react to, such as new data being
  // stored with t.set()
  
  t.lists("all").then(function (lists) {
    document.getElementById('nbLists').textContent = lists.length;
  });

  t.cards("all").then(function (cards) {
    document.getElementById('nbCards').textContent = cards.length;
  });

});
