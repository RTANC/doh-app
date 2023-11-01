<template>
  <v-container fluid>
    <v-card>
      <v-card-title primary-title>
        <v-container fluid>
          <v-row>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <input ref="file" type="file" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" @change="handleFileUpload" v-show="false">
            </v-col>
            <v-col cols="12" xs="3" sm="3" md="3" lg="3" class="d-flex justify-end">
              <v-btn color="info" @click="uploadExcel" dark> <v-icon left>mdi-upload</v-icon> Upload Excel</v-btn>
            </v-col>
            <v-col cols="12" xs="3" sm="3" md="3" lg="3" class="d-flex justify-end">
              <v-btn color="green darken-4" @click="exportToExcel" dark><v-icon left>mdi-download</v-icon> EXPORT to Excel</v-btn>
            </v-col>
          </v-row>
        </v-container>
        <span>8 &nbsp;</span>
        <span class="text-decoration-underline">ASPHALT CONCRETE LEVELING COURSE </span>
        <v-spacer></v-spacer>
        <span class="text-decoration-underline">4 </span>
        <span class="text-decoration-underline">CM. THICK</span>
      </v-card-title>
      <v-card-text>
        <v-form ref="form" v-model="valid">
        <v-container fluid>
          <v-row>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <p class="text-body-1 font-weight-medium black--text">ปริมาณงาน ASPHALT CONCRETE ทั้งโครงการ</p>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between align-baseline">
                <v-icon>mdi-equal</v-icon>
                <span style="width: 80%;">
                  <v-text-field v-model="asphaltConcrete" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)" autofocus></v-text-field>
                </span>
                <span class="text-body-1 black--text font-weight-medium">ตัน</span>
              </div>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <p class="text-body-1 font-weight-medium black--text">ค่ายาง AC</p>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <v-text-field v-model="ac[0]" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(1)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <p class="text-body-1 font-weight-medium black--text d-flex justify-center">ตัน @</p>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="3" lg="3">
              <v-text-field v-model="ac[1]" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(2)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between">
                <v-icon>mdi-equal</v-icon>
                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 1px solid gray; width: 70%;">{{ numberFormatter(calcAc) }}</span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
            <!-- <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <v-text-field prepend-icon="mdi-equal" dense suffix="บาท/ตัน"></v-text-field>
            </v-col> -->
          </v-row>
          <v-row>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <p class="text-body-1 font-weight-medium black--text">ค่าหิน</p>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <v-text-field v-model="rock[0]" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(3)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <p class="text-body-1 font-weight-medium black--text d-flex justify-center">ลบ.ม. @</p>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="3" lg="3">
              <v-text-field v-model="rock[1]" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(4)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between">
                <v-icon>mdi-equal</v-icon>
                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 1px solid gray; width: 70%;">{{ numberFormatter(calcRock) }}</span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
            <!-- <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <v-text-field prepend-icon="mdi-equal" dense suffix="บาท/ตัน"></v-text-field>
            </v-col> -->
          </v-row>
          <v-row>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <p class="text-body-1 font-weight-medium black--text">ค่าดำเนินการ + ค่าเสื่อมผสมวัสดุแอสพัลท์คอนกรีต</p>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between align-baseline">
                <v-icon>mdi-equal</v-icon>
                <span style="width: 70%;">
                  <v-text-field v-model="processPrice1" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(5)"></v-text-field>
                </span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="1" sm="1" md="1" lg="1">
              <p class="text-body-1 font-weight-medium black--text">ค่าขนส่ง</p>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="2" lg="2">
              <v-text-field v-model="transport[0]" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(6)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="3" lg="3">
              <p class="text-body-1 font-weight-medium d-flex justify-center black--text">ก.ม. ( 1 ใน 4 ของระยะทางของโครงการ)</p>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between align-baseline">
                <v-icon>mdi-equal</v-icon>
                <span style="width: 70%;">
                  <v-text-field v-model="transport[1]" type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(7)"></v-text-field>
                </span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="4" sm="4" md="4" lg="4">
              <p class="text-body-1 font-weight-medium black--text">ค่าดำเนินการ + ค่าเสื่อมปูลาดและบดทับหนา</p>
            </v-col>
            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
              <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 1px solid gray;">{{ thickness }}</span>
              <!-- <v-text-field dense reverse value="4"></v-text-field> -->
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="3" lg="3">
              <p class="text-body-1 font-weight-medium black--text">ซม. (บนผิวแทคโค้ด)</p>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="1" sm="1" md="2" lg="2">
              <v-text-field v-model="processPrice2[0]" type="number" filled dense reverse hide-spin-buttons prepend-icon="mdi-equal" @keydown.enter="focusNextInput(8)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="2" lg="2">
              <v-text-field v-model="processPrice2[1]" type="number" filled dense reverse hide-spin-buttons prepend-icon="mdi-close" @keydown.enter="focusNextInput(9)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="1" sm="1" md="2" lg="2">
              <v-text-field v-model="processPrice2[2]" type="number" filled dense reverse hide-spin-buttons prepend-icon="mdi-close" @keydown.enter="focusNextInput(10)"></v-text-field>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between align-baseline">
                <v-icon>mdi-equal</v-icon>
                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 1px solid gray; width: 70%;">{{ numberFormatter(calcProcess2) }}</span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
            <!-- <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <v-text-field prepend-icon="mdi-equal" dense suffix="บาท/ตัน"></v-text-field>
            </v-col> -->
          </v-row>
          <v-row>
            <v-col cols="12" xs="1" sm="1" md="6" lg="6">
              <p class="text-body-1 font-weight-medium black--text">ค่าใช้จ่ายรวม</p>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between">
                <v-icon>mdi-equal</v-icon>
                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 3px double gray; width: 70%;">{{ numberFormatter(calcTotal) }}</span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
            <!-- <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <v-text-field prepend-icon="mdi-equal" dense suffix="บาท/ตัน" readonly></v-text-field>
            </v-col> -->
          </v-row>
          <v-row>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <p class="text-body-1 font-weight-medium black--text d-flex justify-end">ค่างานต้นทุนที่ใช้</p>
            </v-col>
            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <div class="d-flex justify-space-between">
                <v-icon>mdi-equal</v-icon>
                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 3px double gray; width: 70%;">{{ numberFormatter(calcTotal) }}</span>
                <span class="text-body-1 black--text font-weight-medium">บาท/ตัน</span>
              </div>
            </v-col>
            <!-- <v-col cols="12" xs="6" sm="6" md="6" lg="6">
              <v-text-field prepend-icon="mdi-equal" dense suffix="บาท/ตัน" readonly></v-text-field>
            </v-col> -->
          </v-row>
        </v-container>
        </v-form>
      </v-card-text>
    </v-card>
  </v-container>
</template>
<script>
import XLSX from 'xlsx'
export default {
  name: 'AboutView',
  data: () => ({
    valid: false,
    asphaltConcrete: 0,
    ac: [0, 0],
    rock: [0, 0],
    transport: [0, 0],
    processPrice1: 0,
    thickness: 4,
    processPrice2: [0, 0, 0]
  }),
  methods: {
    numberFormatter (v) {
      return new Intl.NumberFormat('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(v)
    },
    focusNextInput (index) {
      if (index < this.$refs.form.inputs.length - 1) {
        this.$refs.form.inputs[index + 1].reset()
        this.$refs.form.inputs[index + 1].focus()
      } else {
        this.$refs.form.inputs[0].focus()
      }
    },
    uploadExcel () {
      this.$refs.file.click()
    },
    async handleFileUpload (e) {
      try {
        const wb = XLSX.read(await e.target.files[0].arrayBuffer(), { cellNF: true })
        const ws = wb.Sheets[wb.SheetNames[0]]
        if (!ws || !ws['!ref']) return
        // eslint-disable-next-line
        this.asphaltConcrete = parseFloat(ws['H2'].v)
        // eslint-disable-next-line
        this.ac = [parseFloat(ws['D3'].v), parseFloat(ws['F3'].v)]
        // eslint-disable-next-line
        this.rock = [parseFloat(ws['D4'].v), parseFloat(ws['F4'].v)]
        // eslint-disable-next-line
        this.processPrice1 = parseFloat(ws['H5'].v)
        // eslint-disable-next-line
        this.transport = [parseFloat(ws['D6'].v), parseFloat(ws['H6'].v)]
        // eslint-disable-next-line
        this.thickness = parseInt(ws['F7'].v)
        // eslint-disable-next-line
        this.processPrice2 = [parseFloat(ws['B8'].v), parseFloat(ws['D8'].v), parseFloat(ws['F8'].v)]
      } catch (error) {
        console.log(error)
      }
    },
    exportToExcel () {
      // eslint-disable-next-line quotes
      const table = `<table border="1" cellspacing="0" cellpadding="3">
        <tbody>
            <tr>
                <td colspan="7">8 &nbsp; ASPHALT CONCRETE LEVELING COURSE</td>
                <td colspan="2">4CM. THICK</td>
            </tr>
            <tr>
                <td colspan="6">ปริมาณงาน ASPHALT CONCRETE ทั้งโครงการ</td>
                <td>=</td>
                <td>${this.asphaltConcrete}</td>
                <td>ต้น</td>
            </tr>
            <tr>
                <td colspan="3">ค่ายาง AC</td>
                <td>${this.ac[0]}</td>
                <td>ตัน @</td>
                <td>${this.ac[1]}</td>
                <td>=</td>
                <td>${this.numberFormatter(this.calcAc)}</td>
                <td>บาท/ตัน</td>
            </tr>
            <tr>
                <td colspan="3">ค่าหิน</td>
                <td>${this.rock[0]}</td>
                <td>ลบ.ม. @</td>
                <td>${this.rock[1]}</td>
                <td>=</td>
                <td>${this.numberFormatter(this.calcRock)}</td>
                <td>บาท/ตัน</td>
            </tr>
            <tr>
                <td colspan="6">ค่าดำเนินการ + ค่าเสื่อมผสมวัสดุแอสพัลท์คอนกรีต</td>
                <td>=</td>
                <td>${this.processPrice1}</td>
                <td>บาท/ตัน</td>
            </tr>
            <tr>
                <td colspan="3">ค่าขนส่ง</td>
                <td>${this.transport[0]}</td>
                <td colspan="2">ก.ม. ( 1 ใน 4 ของระยะทางของโครงการ)</td>
                <td>=</td>
                <td>${this.transport[1]}</td>
                <td>บาท/ตัน</td>
            </tr>
            <tr>
                <td colspan="5">ค่าดำเนินการ + ค่าเสื่อมปูลาดและบดทับหนา</td>
                <td>${this.thickness}</td>
                <td colspan="3">ซม. (บนผิวแทคโค้ด)</td>
            </tr>
            <tr>
                <td>=</td>
                <td>${this.processPrice2[0]}</td>
                <td>x</td>
                <td>${this.processPrice2[1]}</td>
                <td>x</td>
                <td>${this.processPrice2[2]}</td>
                <td>=</td>
                <td>${this.numberFormatter(this.calcProcess2)}</td>
                <td>บาท/ตัน</td>
            </tr>
            <tr>
                <td colspan="6">ค่าใช้จ่ายรวม</td>
                <td>=</td>
                <td>${this.numberFormatter(this.calcTotal)}</td>
                <td>บาท/ตัน</td>
            </tr>
            <tr>
                <td colspan="6">ค่างานต้นทุนที่ใช้</td>
                <td>=</td>
                <td>${this.numberFormatter(this.calcTotal)}</td>
                <td>บาท/ตัน</td>
            </tr>
        </tbody>
    </table>`
      const wb = XLSX.read(table, { type: 'string' })
      XLSX.writeFile(wb, 'ASPHALT CONCRETE LEVELING COURSE.xlsx', { raw: true })
    }
  },
  computed: {
    calcAc: function () {
      return parseFloat(this.ac[0] * this.ac[1])
    },
    calcRock: function () {
      return parseFloat(this.rock[0] * this.rock[1])
    },
    calcTransport: function () {
      return parseFloat(this.transport[1])
    },
    calcProcess1: function () {
      return parseFloat(this.processPrice1)
    },
    calcProcess2: function () {
      return parseFloat(this.processPrice2[0] * this.processPrice2[1] * this.processPrice2[2])
    },
    calcTotal: function () {
      return (this.calcAc + this.calcRock + this.calcProcess1 + this.calcTransport + this.calcProcess2)
    }
  }
}
</script>
