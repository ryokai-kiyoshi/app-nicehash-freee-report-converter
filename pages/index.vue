<template>
  <div>
        <div class="jumbotron jumbotron-fluid">
            <div class="container pghelper-bs4-grid-parent-relative">
                <h1 class="display-4">nicehash マイニング実績レポート変換</h1>
                <p class="lead">nicehashのマイニング実績をFreeeの入金履歴データ(XLSX)に変換します</p>
            </div>
        </div>
        <div class="container">
            <div class="row">
                <div class="col">
                    <div class="bg-light card mb-3">
                        <div class="card-header">nicehash支払い実績データ</div>
                        <div class="card-body">
                            <div class="container" style="height: 200px; border-style: dotted; " 
                              @dragover.prevent.stop="dragAction" 
                              @drop.prevent.stop="dropNicehashReport">
                                <div class="row justify-content-center">
                                    <div class="col-md-4">
                                        <div style="text-align: center;">
                                            ここにファイルをドロップ
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="row justify-content-center">
                <div class="col-md-4" style="text-align: center;">
                    <button type="button rounded-pill" class="btn btn-dark">generate</button>
                </div>
            </div>
        </div>
  </div>
</template>

<script lang="ts">
import { Vue, Component } from "nuxt-property-decorator"
import Moment from 'moment'
import Papa from 'papaparse'
import XLSX, { WritingOptions } from 'xlsx'

@Component({})
export default class Index extends Vue {

  buildReport( csv:any) : any[]{
    const data = csv.data;

    console.log("lines: "+data.length)

    let reportList = []
    for(let i = 1; i < data.length; ++i ){

      const item = data[i]
      if(item[0] == ''){
        break
      }
      const datetime = Moment.utc(item[0].substring(0,19))
      const purpose = item[2]
      const amountBtc = parseFloat(item[3])
      const exchange_rate = parseFloat(item[4])
      const amountJpy = parseFloat(item[5])

      const line = {
        datetime: datetime,
        purpose: purpose,
        amountBtc: amountBtc,
        exchangeRate: exchange_rate,
        amountJpy: amountJpy
      }

      console.log( "NH["+i+"]"+JSON.stringify(line))

      reportList.push(line)
    }

    return(reportList)
  }

  convertFreeeIncomeReport(nicehash:any[]) :any[] {
    let freee = []

    for(let i = 0; i < nicehash.length; ++i){
      const nhItem = nicehash[i]

      // filter
      if(nhItem.purpose == 'Legacy transfer' || nhItem.purpose == 'Repayment' || nhItem.purpose == 'Hashpower mining' ) {

        const fDate = new Date(nhItem.datetime.format('YYYY-MM-DD'))
        const fCategory = '収入'
        const fAccount = '売上高'
        const fAmount = nhItem.amountJpy
        const fTaxType = '課税売上10%'
        const fDueDate = ''
        const fSettlementAccount = 'nicehash'
        const fCustomer = 'nicehash'
        const fItem = 'BTCマイニング'
        const fDepartment = '仮想通貨'
        const fMemo = ''
        const fRemarks = 'BTC: '+nhItem.amountBtc+', BTC/JPY: '+nhItem.exchangeRate+', op:'+nhItem.purpose

        const freeeItem = {
          date: fDate, 
          category: fCategory,
          account: fAccount,
          amount: fAmount,
          taxType: fTaxType,
          dueDate: fDueDate,
          settlementAccount: fSettlementAccount,
          customer: fCustomer,
          item: fItem,
          department: fDepartment,
          memo: fMemo,
          remarks: fRemarks
        }

        console.log("freee["+i+"]:"+JSON.stringify(freeeItem))
        freee.push(freeeItem)
      }
    }
    return(freee)
  }

  buildFreeeIncomeReport(freee: any[]){
    let slips = [['発生日',	'収支区分',	'勘定科目',	'金額',	'税区分',	'決済期日',	'決済口座',	'取引先',	'品目',	'部門',	'メモタグ',	'備考' ]];

    
    for( let i = 0; i < freee.length; ++i){

      const item = freee[i]

      const row = [item.date, item.category, item.account, item.amount, item.taxType,
       item.dueDate, item.settlementAccount, item.customer, item.item, item.department, item.memo, item.remarks
      ]
      slips.push(row)
    }


    const wb = XLSX.utils.book_new()
   /* convert an array of arrays in JS to a CSF spreadsheet */
    const ws = XLSX.utils.aoa_to_sheet(slips, {cellDates:false});
    XLSX.utils.book_append_sheet(wb, ws, "収入取引データ");

    // 新しく作成するエクセルファイルの作成オプションを設定します。
    const opts : WritingOptions = {
      bookType: "xlsx",
      bookSST: false,
      type: 'array',
      compression: true
    }
    
    // 上記オプションを使って Blobオブジェクトに出力します。
    const blob = new Blob(
      [XLSX.write(wb, opts)],
      {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}
    );

    this.downloadFile(blob, 'test.xlsx')
  }
  public async parseNicehashReport(file : any) : Promise<any[]> {

    const csvBin : Uint8Array = await this.readBinFileSync(file)
    const csvString : string = (new TextDecoder).decode(csvBin)

    // console.log(csvString)

    const csv = Papa.parse( csvString)

    console.log("csv parsed."+JSON.stringify(csv))
    const nicehashReport = this.buildReport(csv)
    console.log("parsed.")

    return nicehashReport
  }

  public dragAction(evt : DragEvent) : void {
    console.log(" drap over.")
    evt.stopPropagation()
    evt.preventDefault()
    if( evt.dataTransfer){
      evt.dataTransfer.dropEffect = 'copy'; // Explicitly show this is a copy.
    }

  }

  public async dropNicehashReport(evt : DragEvent) : Promise<void> {
    console.log(" drop.")
    if(evt.dataTransfer){
      let _files = evt.dataTransfer.files

      if(_files && _files.length > 0){
        const nh = await this.parseNicehashReport(_files[0])
        const freee = this.convertFreeeIncomeReport(nh)
        const income = this.buildFreeeIncomeReport(freee)
      }
    }

  }

  async readBinFileSync(file : File) : Promise<Uint8Array>{
    return new Promise((resolve, reject) => {
  
      console.log("readFileSycn:here")
      let reader = new FileReader();
  
      reader.onload = (e:any) => {
        if(e.target){
          let data = new Uint8Array(e.target.result)
          resolve(data);
        }

      };
  
      reader.onerror = reject;
  
      reader.readAsArrayBuffer(file);
    })
  }

  downloadFile(blob:Blob, filename:string) : void {
    // ダウンロードさせる
    const url = (window.URL || window.webkitURL).createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url                 // ダウンロード先URLに指定.
    a.download = filename       // ダウンロードファイル名を指定.
    document.body.appendChild(a) // aタグ要素を画面に一時的に追加する（これをしないとFirefoxで動作しない）.
    a.click()                    // クリックすることでダウンロードを開始.
    document.body.removeChild(a) // 不要になったら削除.
   }
}
//export default Vue.extend({})
</script>

<style>
/* .container {
  margin: 0 auto;
  min-height: 100vh;
  display: flex;
  justify-content: center;
  align-items: center;
  text-align: center;
}

.title {
  font-family:
    'Quicksand',
    'Source Sans Pro',
    -apple-system,
    BlinkMacSystemFont,
    'Segoe UI',
    Roboto,
    'Helvetica Neue',
    Arial,
    sans-serif;
  display: block;
  font-weight: 300;
  font-size: 100px;
  color: #35495e;
  letter-spacing: 1px;
}

.subtitle {
  font-weight: 300;
  font-size: 42px;
  color: #526488;
  word-spacing: 5px;
  padding-bottom: 15px;
}

.links {
  padding-top: 15px;
} */
</style>
