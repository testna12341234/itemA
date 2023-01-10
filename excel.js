let selectedFile;
console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let data = [{
    "name": "jayanth",
    "data": "scd",
    "abc": "sdef"
}]

const groupBy = (array, key) => {
    return array.reduce((result, currentValue) => {
        (result[currentValue[key]] = result[currentValue[key]] || []).push(currentValue);
        return result;
    }, {});
};


function generateAsExcel(data) {
    try {

        const workbook = XLSX.utils.book_new();

        for (let key in data) {
            const worksheet = XLSX.utils.json_to_sheet(data[key]);
            XLSX.utils.book_append_sheet(workbook, worksheet, key);
        }

        let res = XLSX.write(workbook, {
            type: "array"
        });
        console.log(`${res.byteLength} bytes generated`);
    } catch (err) {
        console.log("Error:", err);
    }
}



function compare(a, b) {
    // Use toUpperCase() to ignore character casing
    const bandA = a.收件人電話號碼.toUpperCase();
    const bandB = b.收件人電話號碼.toUpperCase();

    let comparison = 0;
    if (bandA > bandB) {
        comparison = 1;
    } else if (bandA < bandB) {
        comparison = -1;
    }
    return comparison;
}


function outputAddress(address) {
    let cNw = ["石塘咀", "堅尼地城", "西營盤", "上環", "中環", "金鐘", "半山區", "山頂 ", "Kennedy Town", "Shek Tong Tsui", "Sai Ying Pun", "Sheung Wan", "Central", "Admiralty", "Mid-levels", "Peak"]
    let wc = ["灣仔", "銅鑼灣", "跑馬地", "大坑", "掃桿埔", "渣甸山", "Wan Chai", "Causeway Bay", "Happy Valley", "Tai Hang", "So Kon Po", "Jardine"]
    let eastern = ["天后", "寶馬山", "北角", "鰂魚涌", "西灣河", "筲箕灣", "柴灣", "小西灣", "Tin Hau", "Braemar Hill", "North Point", "Quarry Bay", "Sai Wan Ho", "Shau Kei Wan", "Chai Wan", "Siu Sai Wan"]
    let southern = ["薄扶林", "香港仔", "鴨脷洲", "黃竹坑", "壽臣山", "淺水灣", "舂磡角", "赤柱", "大潭", "石澳", "Pok Fu Lam", "Aberdeen", "Ap Lei Chau", "Wong Chuk Hang", "Shouson Hill", "Repulse Bay", "Chung Hom Kok", "Stanley", "Tai Tam", "Shek O"]
    let tsm = ["尖沙咀", "油麻地", "西九龍填海區", "京士柏", "旺角", "大角咀", "Tsim Sha Tsui", "Yau Ma Tei", "West Kowloon Reclamation", "King\'s Park", "Mong Kok", "Tai Kok Tsui"]
    let sss = ["美孚", "荔枝角", "長沙灣", "深水埗", "石硤尾", "又一村", "大窩坪", "昂船洲", "Mei Foo", "Lai Chi Kok", "Cheung Sha Wan", "Sham Shui Po", "Shek Kip Mei", "Yau Yat Tsuen,Tai Wo Ping", "Stonecutters Island"]
    let kwc = ["紅磡", "土瓜灣", "馬頭角", "馬頭圍", "啟德", "九龍城", "何文田", "九龍塘", "筆架山", "Hung Hom", "To Kwa Wan", "Ma Tau Kok", "Ma Tau Wai", "Kai Tak", "Kowloon City", "Ho Man Tin", "Kowloon Tong", "Beacon Hill"]
    let wts = ["新蒲崗", "黃大仙", "東頭", "橫頭磡", "樂富", "鑽石山", "慈雲山", "牛池灣", "San Po Kong", "Wong Tai Sin", "Tung Tau", "Wang Tau Hom", "Lok Fu", "Diamond Hill", "Tsz Wan Shan", "Ngau Chi Wan"]
    let kt = ["坪石", "九龍灣", "牛頭角", "佐敦谷", "觀塘", "秀茂坪", "藍田", "油塘、 鯉魚門", "Ping Shek", "Kowloon Bay", "Ngau Tau Kok", "Jordan Valley", "Kwun Tong", "Sau Mau Ping", "Lam Tin", "Yau Tong", "Lei Yue Mun"]
    let ktsing = ["葵涌", "青衣", "Kwai Chung", "Tsing Yi"]
    let tw = ["荃灣", "梨木樹", "汀九", "深井", "青龍頭", "馬灣", "欣澳", "Tsuen Wan", "Lei Muk Shue", "Ting Kau", "Sham Tseng", "Tsing Lung Tau", "Ma Wan", "Sunny Bay"]
    let tm = ["大欖涌", "掃管笏", "屯門", "藍地", "Tai Lam Chung", "So Kwun Wat", "Tuen Mun", "Lam Tei"]
    let yl = ["洪水橋", "廈村", "流浮山", "天水圍", "元朗", "新田", "落馬洲", "錦田", "石崗", "八鄉", "Hung Shui Kiu", "Ha Tsuen", "Lau Fau Shan", "Tin Shui Wai", "Yuen Long", "San Tin", "Lok Ma Chau", "Kam Tin", "Shek Kong", "Pat Heung"]
    let north = ["粉嶺", "聯和墟", "上水", "石湖墟", "沙頭角", "鹿頸", "烏蛟騰", "Fanling", "Luen Wo Hui", "Sheung Shui", "Shek Wu Hui", "Sha Tau Kok", "Luk Keng", "Wu Kau Tang"]
    let tp = ["大埔墟", "大埔", "大埔滘", "大尾篤", "船灣", "樟木頭", "企嶺下", "Tai Po Market", "Tai Po", "Tai Po Kau", "Tai Mei Tuk", "Shuen Wan", "Cheung Muk Tau", "Kei Ling Ha"]
    let st = ["大圍", "沙田", "火炭", "馬料水", "烏溪沙", "馬鞍山", "Tai Wai", "Sha Tin", "Fo Tan", "Ma Liu Shui", "Wu Kai Sha", "Ma On Shan"]
    let sk = ["清水灣", "西貢", "大網仔", "將軍澳", "坑口", "調景嶺", "馬游塘", "Clear Water Bay", "Sai Kung", "Tai Mong Tsai", "Tseung Kwan O", "Hang Hau", "Tiu Keng Leng", "Ma Yau Tong"]
    let islands = ["長洲", "坪洲", "大嶼山", "東涌", "南丫島", "Cheung Chau", "Peng Chau", "Lantau Island", "Lamma Island"]


    let city = [];
    cNw.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("中西區");
            return city;
        }
    });
    wc.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("灣仔");
            return city;
        }
    });
    eastern.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("東區");
            return city;
        }
    });
    southern.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("南區");
            return city;
        }
    });
    tsm.forEach(a => {
        if (address.includes(a)) {
            city.push("九龍");
            city.push("油尖旺");
            return city;
        }
    });
    sss.forEach(a => {
        if (address.includes(a)) {
            city.push("九龍");
            city.push("深水埗");
            return city;
        }
    });
    kwc.forEach(a => {
        if (address.includes(a)) {
            city.push("九龍");
            city.push("九龍城");
            return city;
        }
    });
    wts.forEach(a => {
        if (address.includes(a)) {
            city.push("九龍");
            city.push("黃大仙");
            return city;
        }
    });
    kt.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("觀塘");
            return city;
        }
    });
    ktsing.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("葵青");
            return city;
        }
    });
    tw.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("荃灣");
            return city;
        }
    });
    tm.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("屯門");
            return city;
        }
    });
    yl.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("元朗");
            return city;
        }
    });
    north.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("北區");
            return city;
        }
    });
    tp.forEach(a => {
        if (address.includes(a)) {
            city.push("新界");
            city.push("大埔");
            return city;
        }
    });
    st.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("沙田");
            return city;
        }
    });
    sk.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("西貢");
            return city;
        }
    });
    islands.forEach(a => {
        if (address.includes(a)) {
            city.push("港島");
            city.push("離島");
            return city;
        }
    });

    return city;
}

function padWithLeadingZeros(num, totalLength) {
    return String(num).padStart(totalLength, '0');
}

function countTotal(list, number) {
    var total =0 ;
    list.forEach(a => {
        if (a.收件人電話號碼 == number) {
            total++;
            console.log(a.訂單號碼)
        } 
    });
    return total;
}

/*function renameKey ( obj, oldKey, newKey ) {
  obj[newKey] = obj[oldKey];
  delete obj[oldKey];
}*/

function dataProcess(list) {
    // let address= a.地址1;



    let count = 0;
    let lastNumber = 0;
    let newList = [];
    let total = 1;

    list.forEach(a => {
         total = countTotal(list, a.收件人電話號碼);

            if (lastNumber != 0&& a.收件人電話號碼 != lastNumber) {
                count++;
    
            }
        
        lastNumber = a.收件人電話號碼;
        let order = "VFZ#"
        let number = padWithLeadingZeros(count, 3); //001
        order += number + "(";
        order += total + ")";
        let city = outputAddress(a.地址);



        let b = {
            "订单号": order,
            "客户代码": "OHSB",
            "客户名称": "",
            "订单时间": "",
            "快递单号": "",
            "库存类型": "銷售",
            "关联单号": a.訂單號碼,
            "发货公司": "",
            "发货联系人": "OHBABY STAR BLESS COMPANY LIMITED",
            "发货手机": "",
            "发货电话": "",
            "发货国家": "",
            "发货省": "",
            "发货市": "",
            "发货区": "",
            "收货公司": "",
            "收货联系人": a.收件人,
            "收货手机": a.收件人電話號碼,
            "收货电话": "",
            "收货国家": "",
            "收货省": "香港",
            "收货市": city[0],
            "收货区": city[1],
            "收货详细地址": a.地址,
            "收货邮政编码": "",
            "收货电子邮箱": "",
            "收货身份证号": "",
            "备注": "",
            "商品编码": "",
            "条码": "",
            "商品名称": a.商品名稱,
            "订单数量": a.數量,
            "单价": "",
            "入仓订单号": "",
            "批次号": "",
            "生产日期": "",
            "过期日期": ""

        }
        newList.push(b);
    });
    return newList;
}


document.getElementById('button').addEventListener("click", () => {
    XLSX.utils.json_to_sheet(data, 'out.xlsx');
    if (selectedFile) {
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event) => {
            let data = event.target.result;
            let workbook = XLSX.read(data, {
                type: "binary"
            });
            let ans = [];
            console.log(workbook);
            workbook.SheetNames.forEach(sheet => {
                let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
                ans.push(rowObject);

            });

            let allList = [];

            let ftList = [];
            let kfList = [];
            let finalList = [];
            let sfShopList = [];
            const result = groupBy(ans[0], '送貨方式');
            //  var obj = JSON.stringify(result, null, 2);

            if(Object.keys(result).includes("$35 順豐寄付配送 (運費需先入數)")){
                result["$35 順豐寄付配送 (運費需先入數)"].forEach(function(order) {
                    finalList.push(order)
             });
            }
             if(Object.keys(result).includes("(只限加單)順豐寄付派送(只適用於已有訂單並已付運費)")){
            result["(只限加單)順豐寄付派送(只適用於已有訂單並已付運費)"].forEach(function(order) {
                finalList.push(order)
            });
        }
 if(Object.keys(result).includes("順豐到付(工商/住宅)")){
             result["順豐到付(工商/住宅)"].forEach(function(order) {
                finalList.push(order)
            });
         }
          if(Object.keys(result).includes("順豐到付(順便智能櫃取件)")){
              result["順豐到付(順便智能櫃取件)"].forEach(function(order) {
                finalList.push(order)
            });
          }
           if(Object.keys(result).includes("順豐到付(順豐營業點取件)")){
               result["順豐到付(順豐營業點取件)"].forEach(function(order) {
                finalList.push(order)
            });
           }
            if(Object.keys(result).includes("順豐到付 (順豐站取件)")){
                result["順豐到付 (順豐站取件)"].forEach(function(order) {
                finalList.push(order)
            });
            }
             if(Object.keys(result).includes("順豐到付 OK便利店取件 (經順豐速運) 重量限制係 5公斤或以下，最大體積36*30*25厘米")){
                 result["順豐到付 OK便利店取件 (經順豐速運) 重量限制係 5公斤或以下，最大體積36*30*25厘米"].forEach(function(order) {
                finalList.push(order)
            });
             }
              if(Object.keys(result).includes("順豐到付7-11便利店取件 (經順豐速運) 重量限制係 5公斤或以下，最大體積36*30*25厘米")){
                  result["順豐到付7-11便利店取件 (經順豐速運) 重量限制係 5公斤或以下，最大體積36*30*25厘米"].forEach(function(order) {
                finalList.push(order)
            });
              }

            allList.push(finalList);

            //  ftList= groupBy(ftList, '收件人電話號碼');


            allList.forEach(e => {
                e.sort(compare)
            })

            let test = [];
             test = dataProcess(allList[0]);



            console.log(test);
            document.getElementById("jsondata").innerHTML = JSON.stringify(test, undefined, 4); //JSON.stringify(


            const workBook = XLSX.utils.book_new();

            const workSheet3 = XLSX.utils.json_to_sheet(test);
            XLSX.utils.book_append_sheet(workBook, workSheet3, "順豐到付(順便智能櫃取件)");




            // XLSX.write(workBook, { bookType: "xlsx", type: "buffer" });
             //XLSX.write(workBook, { bookType: "xlsx", type: "binary" });
             //XLSX.writeFile(workBook,"newExcel.xlsx");

            /*    filename='reports.xlsx';     
                var ws = XLSX.utils.json_to_sheet(json);
                var wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, "People");
                XLSX.writeFile(wb,filename);*/
        }
    }
});