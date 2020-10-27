<template>
  <div id="app">
    <link
      rel="stylesheet"
      href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css"
    />

    <div class="container">
           
      <div class="kz">
        <input class="sub" type="text" v-model="sub1" /><br /><br />
        <input ref="input" type="file" @change="onChange" />
        <xlsx-read :file="file">
          <xlsx-json @parsed="parsed"></xlsx-json>
        </xlsx-read>
        <div v-model="sub1" v-if="finished" class="finish">
          Рассылка закончена
        </div>
        <p>{{ sub1 }}</p>
      </div>

      <div class="kz">
        <input class="sub" type="text" v-model="sub2" /><br /><br />
        <input ref="input2" type="file" @change="onChange2" />
        <xlsx-read :file="file2">
          <xlsx-json @parsed="second"></xlsx-json>
        </xlsx-read>
        <div v-model="sub2" v-if="finished2" class="finish">
          Рассылка закончена
        </div>
        <p>{{ sub2 }}</p>
      </div>
    </div>
  </div>
</template>

<script>
// import emailjs from "emailjs-com"; //_____API
import { XlsxRead, XlsxJson, XlsxTable, XlsxSheets } from "vue-xlsx";

export default {
  name: "app",
  data: () => ({
    file: null,
    file2: null,
    selectedSheet: null,
    collection: [],
    finished: false,
    finished2: false,
    sub1: "ЗП 2020 фг расчет",
    sub2: "КЗ наушники ERGO",
  }),

  async mounted() {
    // console.log(this);
  },

  methods: {
    onChange(e) {
      this.file = e.target.files ? e.target.files[0] : null;
      this.$refs.input.disabled = true;
      this.finished = true;
    },

    onChange2(e) {
      this.file2 = e.target.files ? e.target.files[0] : null;
      this.$refs.input2.disabled = true;
      this.finished2 = true;
    },

    sendEmail: (subj, email, ccemail, textMail) => {
      Email.send({
        Host: "smtp.gmail.com",
        Username: "yugcontract.shevchuk@gmail.com",
        Password: process.env.VUE_APP_PASS,
        To: [email, ccemail],
        From: "yugcontract.shevchuk@gmail.com",
        Subject: subj,
        Body: textMail,
      }).then((message) => console.log(message));
    },

    parsed(e) {
      e.map((data, index) => {
        if (index > 1) {
          let keys = Object.keys(data);
          let body = `
          <h3 style="color: rgb(0, 41, 128)">Добрый день, ${data[keys[2]]} </h3>
          <table style="border-collapse:collapse; width: 480px; font-size: 14px">
          <colgroup width="200"></colgroup>
          <colgroup width="194"></colgroup>
          <colgroup width="115"></colgroup>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" rowspan=2 height="40" align="left" valign=middle>ДС</td>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Поступление ДС</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="2486992.37" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-">  ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[3]])} </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>ХО 5465</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="184134.86" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"> ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[4]])}  </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" colspan=2 height="20" align="left" valign=bottom bgcolor="#DEEBF7"><b>Итого ДС</b></td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom bgcolor="#DEEBF7" sdval="2671127.23" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"><b>  ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[5]])}  </b></td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" rowspan=6 height="120" align="left" valign=middle>Продажи</td>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Моб тел Ерго</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="20841" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-">  ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[6]])}  </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Сетевое</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="14044" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-">  ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[7]])}  </td>
          </tr>
          <tr>
            <td style="width: 200px; padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Носимые устройства Ерго</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="1154.05" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"> ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[8]])}  </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Powerbank</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="37885.93" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"> ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[9]])} </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Наушники</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="21987.68" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"> ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[10]])}  </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="left" valign=bottom>Флеш</td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom sdval="11500" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-">  ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[11]])}  </td>
          </tr>
          <tr>
            <td style="padding-left: 5px; border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" colspan=2 height="20" align="left" valign=bottom bgcolor="#DEEBF7" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"><b> Итого продажи </b></td>
            <td style="border-top: 1px solid #bfbfbf; border-bottom: 1px solid #bfbfbf; border-left: 1px solid #bfbfbf; border-right: 1px solid #bfbfbf" align="center" valign=bottom bgcolor="#DEEBF7" sdval="107412.66" sdnum="1033;0;_-* #,##0.00_-;-* #,##0.00_-;_-* &quot;-&quot;??_-;_-@_-"><b> ${new Intl.NumberFormat(
              "ru-RU",
              { maximumFractionDigits: 2 }
            ).format(data[keys[12]])} </b></td>
          </tr>
        </table>
      `;
          this.sendEmail(this.sub1, data[keys[0]], data[keys[1]], body);
        }
      })
    },

    second(ev) {
      ev.map((d) => {
        console.log(d);
        let body = `
          <h3 style="color: rgb(0, 41, 128)">Добрый день, ${
            d["Названия строк"]
          } </h3>
          <table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=0
          style='width:169.85pt;margin-left:-.15pt;border-collapse:collapse;mso-yfti-tbllook:
          1184;mso-padding-alt:0cm 5.4pt 0cm 5.4pt'>
          <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:15.0pt'>
            <td width=151 nowrap valign=bottom style='width:113.15pt;border:solid windowtext 1.0pt;
            mso-border-alt:solid windowtext .5pt;background:#DDEBF7;padding:0cm 5.4pt 0cm 5.4pt;
            height:15.0pt'>
            <p class=MsoNormal><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
            "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:"Times New Roman";
            color:black;mso-fareast-language:RU;mso-bidi-font-weight:bold'>План АКБ<o:p></o:p></span></p>
            </td>
            <td width=76 nowrap valign=bottom style='width:2.0cm;border:solid windowtext 1.0pt;
            border-left:none;mso-border-top-alt:solid windowtext .5pt;mso-border-bottom-alt:
            solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;padding:
            0cm 5.4pt 0cm 5.4pt;height:15.0pt'>
            <p class=MsoNormal align=center style='text-align:center'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:"Times New Roman";
            color:black;mso-fareast-language:RU'> ${
              d["План АКБ"]
            } <o:p></o:p></span></p>
            </td>
          </tr>
          <tr style='mso-yfti-irow:1;height:15.0pt'>
            <td width=151 nowrap valign=bottom style='width:113.15pt;border:solid windowtext 1.0pt;
            border-top:none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
            solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;background:
            #DDEBF7;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>
            <p class=MsoNormal><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
            "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:"Times New Roman";
            color:black;mso-fareast-language:RU;mso-bidi-font-weight:bold'>Факт АКБ<o:p></o:p></span></p>
            </td>
            <td width=76 nowrap valign=bottom style='width:2.0cm;border-top:none;
            border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
            mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
            padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>
            <p class=MsoNormal align=center style='text-align:center'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:"Times New Roman";
            color:black;mso-fareast-language:RU'> ${
              d["Факт АКБ"]
            } <o:p></o:p></span></p>
            </td>
          </tr>
          <tr style='mso-yfti-irow:2;mso-yfti-lastrow:yes;height:15.0pt'>
            <td width=151 nowrap valign=bottom style='width:113.15pt;border:solid windowtext 1.0pt;
            border-top:none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
            solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;background:
            #DDEBF7;padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>
            <p class=MsoNormal><b><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
            "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:"Times New Roman";
            color:black;mso-fareast-language:RU'>Выполнение %<o:p></o:p></span></b></p>
            </td>
            <td width=76 nowrap valign=bottom style='width:2.0cm;border-top:none;
            border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
            mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
            padding:0cm 5.4pt 0cm 5.4pt;height:15.0pt'>
            <p class=MsoNormal align=center style='text-align:center'><span
            style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
            mso-hansi-font-family:Calibri;mso-bidi-font-family:"Times New Roman";
            color:black;mso-fareast-language:RU'>${new Intl.NumberFormat(
              "ru-RU",
              { style: "percent" }
            ).format(+d["Выполнение %"])} <o:p></o:p></span></p>
            </td>
          </tr>
          </table>
      `;
        this.sendEmail(this.sub2, d["Получатель"], d["Копия"], body);
      });
    },
  },
  components: {
    XlsxRead,
    XlsxJson,
    XlsxTable,
    XlsxSheets,
  },
};
</script>


<style lang="scss">
@import "~materialize-css/dist/css/materialize.min.css";

.container {
  width: 90%;
  padding-top: 30px;
  display: flex;
}

.kz {
  width: 50%;
}
.sub {
  font-size: 22px !important;
  width: 300px !important;
}

.finish {
  padding: 10px 10px 10px 0;
  font-size: 22px;
  color: #fd0606ba;
}
</style>
