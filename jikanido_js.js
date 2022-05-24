let table;

(function(){
    //HTML自体にドラッグ＆ドロップしても何もしないようにする
    $('html').on('dragover', function(e){
     e.preventDefault();
     e.stopPropagation();
    });
    $('html').on('drop', function(e){
     e.preventDefault();
     e.stopPropagation();
    });
})();

(function(){
    //ファイルをエリア内にドラッグしてきた
     $('div.droparea').on('dragover', function(e){
        e.preventDefault();
        e.stopPropagation();
        let p = $(this).find('p');
        $(p).text("ここにドロップ");
    });

    //エリアから離れた（初期状態に戻す）
     $('div.droparea').on('dragleave',function(e){
      e.preventDefault();
      e.stopPropagation();
      let p = $(this).find('p');
      $(p).text("ここにファイルをドラッグ＆ドロップするか、クリックしてファイルを選択");
     });

     //ドロップ
     $('div.droparea').on('drop',function(e){
      e.preventDefault();
      e.stopPropagation();
     
      let p = $(this).find('p');
      $(p).text("アップロード中");
       
      //fileを取得
      let file2 = e.originalEvent.dataTransfer.files[0];
     
      //ajaxファイルアップロード関数を呼ぶ（4で説明)
      readJikanidoFile(file2);
     });
})();

(function(){
    $("#divcd").focusout((e)=> {
        let divcd = $(e.target).val();
        table.clearFilter();
        table.addFilter([
            [
            {field:"移動元組織コード", type:"=", value:divcd},
            {field:"移動先組織コード", type:"=", value:divcd},
            ]
        ]);
    });
})();

function readJikanidoFile(file){
    let reader = new FileReader();
    const divcd = $("#divcd").val();

    reader.onload = () => {
        let jikanidoFile = reader.result;
        let csvmap = doAnalyze(jikanidoFile, divcd);
        // 読み込み完了時の処理
        table = new Tabulator("#example-table", {
            height:500, // set height of table to enable virtual DOM
            data:csvmap, //load initial data into table
            layout:"fitColumns", //fit columns to width of table (optional)
            columns:[ //Define Table Columns
                {title:"移動元組織コード", field:"移動元組織コード", sorter:"string"},
                {title:"移動先組織コード", field:"移動先組織コード", sorter:"string"},
                {title:"時間種別", field:"時間種別", sorter:"string"},
                {title:"移動時間", field:"移動時間", align:"right", sorter:"number", formatterParams:{precision:2}, bottomCalc:"sum", bottomCalcParams:{precision:2}},
                {title:"内容", field:"内容", sorter:"string"},
            ],
            rowClick:function(e, row){ //trigger an alert message when the row is clicked
                let num = row._row.data['移動時間'];
                //input.copyの値をクリップボードにコピーする
                const copyToClipboard = (text) => {
                    navigator.clipboard.writeText(text)
                        .then(
                            success => alert('Copyed' + text),
                            error => alert('Failed to copy')
                        );
                };
                copyToClipboard(num);
            },
            initialSort:[
                {column:"時間種別", dir:"desc"}, //sort by this first
            ]
        });

        table.setFilter([
            [
            {field:"移動元組織コード", type:"=", value:divcd},
            {field:"移動先組織コード", type:"=", value:divcd},
            ]
        ]);

        $("#msgarea").text("解析完了");
    };
    reader.readAsArrayBuffer(file);
}

function analyzeJikanido(){
    let file = $("#jikanido").prop('files')[0];
    readJikanidoFile(file);
}

function doAnalyze(file, divcd){
    const book = XLSX.read(file, {
        type: "array",
        codepage: 932
    });

    sheets = book.SheetNames.map((name) => {
        let json = XLSX.utils.sheet_to_json(book.Sheets[name]);
        return json
    });

    let jikanido_sheet = sheets[0];
    let jikanido_sheet2 = jikanido_sheet.map((row) => {
        return {
            "移動元組織コード": row['移動元組織コード'],
            "移動先組織コード": row['移動先組織コード'],
            "時間種別": row['時間種別'],
            "移動時間": toFJikan(row['移動時間'], row['移動元組織コード'] == divcd),
            "内容": row['内容'],
        }
    });

    return jikanido_sheet2;
}

function toFJikan(val, isMinus) {
    let split = val.split(":");
    ret = parseInt(split[0]);
    switch(split[1]) {
    case "15":
        ret += 0.25;
        break;
    case "30":
        ret += 0.5;
        break;
    case "45":
        ret += 0.75;
    }
    if(isMinus){
        return ret = ret * -1;
    }
    return ret;
}