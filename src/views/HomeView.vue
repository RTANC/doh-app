<template>
  <v-container fluid>
    <v-card>
      <v-card-text>
        <v-container fluid>
          <v-row>
            <v-col cols="12" xs="10" sm="10" md="10" lg="10">
              <p>รายการคำนวณหาปริมาณวัสดุ</p>
            </v-col>
            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
              <p>แผ่นที่ 1/1</p>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="10" sm="10" md="10" lg="10">
              <p>โครงการ :</p>
            </v-col>
            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
              <p>UPDATE</p>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="10" sm="10" md="10" lg="10">
              <p>อ้างถึงแบบมาตรฐานเลขที่ PC-304</p>
            </v-col>
            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
              <p>{{ toDay }}</p>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
              <p>ABUTMENT WITHOUT SIDEWALK</p>
            </v-col>
            <v-col cols="12" xs="3" sm="3" md="3" lg="3">
              <v-select
                :items="[{ text: '25', value: 0}, { text: '30', value: 1}]"
                v-model="ln"
                label="Ln (M.)"
                outlined
              ></v-select>
            </v-col>
            <v-col cols="12" xs="3" sm="3" md="3" lg="3">
              <v-select
                :items="[{ text: '12', value: 0}, { text: '15', value: 1}]"
                v-model="w"
                label="W (M.)"
                outlined
              ></v-select>
            </v-col>
            <v-col cols="12" xs="3" sm="3" md="3" lg="3">
              <v-select
                :items="[{ text: '0', value: 1 }, { text: '5', value: 1.0038 }, { text: '10', value: 1.0154 }, { text: '15', value: 1.0353 }, { text: '20', value: 1.0642 }, { text: '25', value: 1.1034 }, { text: '30', value: 1.1547 }, { text: '35', value: 1.2208 }, { text: '40', value: 1.3054 }, { text: '45', value: 1.4142 }]"
                v-model="skew"
                label="SKEW"
                outlined
              ></v-select>
            </v-col>
            <v-col cols="12"  xs="1" sm="1" md="1" lg="1">
              <p>SHEET NO. 7</p>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12">
              <v-form ref="form" v-model="valid">
              <v-simple-table>
                <template v-slot:default>
                  <thead>
                    <tr>
                      <th class="text-center">
                        ลำดับ
                      </th>
                      <th class="text-center">
                        รายการวัสดุ
                      </th>
                      <th class="text-center">
                        ปริมาณ
                      </th>
                      <th class="text-center">
                        หน่วย
                      </th>
                      <th class="text-center">
                        หมายเหตุ
                      </th>
                      <th class="text-center">
                        ราคาต่อหน่วย
                      </th>
                      <th class="text-center">
                        สรุปราคา
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr>
                      <td></td>
                      <td class="font-weight-black">ฐานราก <v-btn color="info" icon @click="popupFooting"><v-icon>mdi-arrow-top-right-bold-box</v-icon></v-btn></td>
                      <td colspan="5"></td>
                    </tr>
                    <tr>
                      <td class="text-center">1</td>
                      <td>- เสาเข็ม 0.525x0.525 ม. SAFETY LOAD 50 TON/PILE</td>
                      <td class="text-center">{{ footingt38[ln][w] }}</td>
                      <td class="text-center">ต้น</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field autofocus v-model="pricePerUnits[0]" outlined dense reverse @keydown.enter="focusNextInput(0)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((footingt38[ln][w] * pricePerUnits[0])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">2</td>
                      <td>- ปริมาณคอนกรีตหยาบ</td>
                      <td class="text-center">{{ numberFormatter(calcLeanConcrete) }}</td>
                      <td class="text-center">ลบม.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[1]" outlined dense reverse @keydown.enter="focusNextInput(1)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcLeanConcrete * pricePerUnits[1])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">3</td>
                      <td>- ปริมาณทรายบดอัด</td>
                      <td class="text-center">{{ numberFormatter(calcCompactSand) }}</td>
                      <td class="text-center">ลบม.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[2]" outlined dense reverse @keydown.enter="focusNextInput(2)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcCompactSand * pricePerUnits[2])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">4</td>
                      <td>- ไม้แบบ</td>
                      <td class="text-center">{{ numberFormatter(calcFormWork) }}</td>
                      <td class="text-center">ตรม.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[3]" outlined dense reverse @keydown.enter="focusNextInput(3)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcFormWork * pricePerUnits[3])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">5</td>
                      <td>- คอนกรีต Strength 35 Mpa.</td>
                      <td class="text-center">{{ numberFormatter(calcConcrete35) }}</td>
                      <td class="text-center">ลบม.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[4]" outlined dense reverse @keydown.enter="focusNextInput(4)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcConcrete35 * pricePerUnits[4])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">6</td>
                      <td class="font-weight-black">- เหล็กเสริม</td>
                      <td class="text-center"></td>
                      <td class="text-center"></td>
                      <td></td>
                      <td class="text-right"></td>
                      <td class="text-right"></td>
                    </tr>
                    <tr>
                      <td class="text-center"></td>
                      <td>  - DB20</td>
                      <td class="text-center">{{ numberFormatter(calcDB20) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[5]" outlined dense reverse @keydown.enter="focusNextInput(5)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcDB20 * pricePerUnits[5])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center"></td>
                      <td>  - DB25</td>
                      <td class="text-center">{{ numberFormatter(calcDB25) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[6]" outlined dense reverse @keydown.enter="focusNextInput(6)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcDB25 * pricePerUnits[6])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">7</td>
                      <td>- ลวดผูกเหล็ก</td>
                      <td class="text-center">{{ numberFormatter(((calcDB20 + calcDB25) * 0.025)) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[7]" outlined dense reverse @keydown.enter="focusNextInput(7)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((((calcDB20 + calcDB25) * 0.025) * pricePerUnits[7])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center"></td>
                      <td class="font-weight-black">ABUTMENT <v-btn color="info" icon @click="popupAbutment"><v-icon>mdi-arrow-top-right-bold-box</v-icon></v-btn></td>
                      <td class="text-center"></td>
                      <td class="text-center"></td>
                      <td></td>
                      <td class="text-right"></td>
                      <td class="text-right"></td>
                    </tr>
                    <tr>
                      <td class="text-center">1</td>
                      <td>- ไม้แบบ</td>
                      <td class="text-center">{{ numberFormatter(calcFormWork2) }}</td>
                      <td class="text-center">ตรม.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[8]" outlined dense reverse @keydown.enter="focusNextInput(8)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcFormWork2 * pricePerUnits[8])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">2</td>
                      <td>- คอนกรีต Strength 35 Mpa.</td>
                      <td class="text-center">{{ numberFormatter(calcConcrete352) }}</td>
                      <td class="text-center">ลบม.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[9]" outlined dense reverse @keydown.enter="focusNextInput(9)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcConcrete352 * pricePerUnits[9])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">3</td>
                      <td class="font-weight-black">- เหล็กเสริม</td>
                      <td class="text-center"></td>
                      <td class="text-center"></td>
                      <td></td>
                      <td class="text-right"></td>
                      <td class="text-right"></td>
                    </tr>
                    <tr>
                      <td class="text-center"></td>
                      <td>  - DB12</td>
                      <td class="text-center">{{ numberFormatter(calcDB12) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[10]" outlined dense reverse @keydown.enter="focusNextInput(10)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcDB12 * pricePerUnits[10])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center"></td>
                      <td>  - DB16</td>
                      <td class="text-center">{{ numberFormatter(calcDB16) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[11]" outlined dense reverse @keydown.enter="focusNextInput(11)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcDB16 * pricePerUnits[11])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center"></td>
                      <td>  - DB20</td>
                      <td class="text-center">{{ numberFormatter(calcDB202) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[12]" outlined dense reverse @keydown.enter="focusNextInput(12)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((calcDB202 * pricePerUnits[12])) }}</td>
                    </tr>
                    <tr>
                      <td class="text-center">4</td>
                      <td>- ลวดผูกเหล็ก</td>
                      <td class="text-center">{{ numberFormatter((calcDB12 + calcDB16 + calcDB202) * 0.025) }}</td>
                      <td class="text-center">กก.</td>
                      <td></td>
                      <td class="d-flex">
                        <v-text-field v-model="pricePerUnits[13]" outlined dense reverse @keydown.enter="focusNextInput(13)" type="number" hide-spin-buttons></v-text-field>
                      </td>
                      <td class="text-right">{{ numberFormatter((((calcDB12 + calcDB16 + calcDB202) * 0.025) * pricePerUnits[13])) }}</td>
                    </tr>
                  </tbody>
                  <tfoot>
                    <tr>
                      <td colspan="5"></td>
                      <td class="text-right">รวมเป็นราคา</td>
                      <td class="text-right"></td>
                    </tr>
                  </tfoot>
                </template>
              </v-simple-table>
              </v-form>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" class="d-flex justify-center">
              <p class="text-decoration-underline subtitle-2">ภาพประกอบ</p>
            </v-col>
          </v-row>
          <v-row>
            <v-col cols="12" class="d-flex justify-center">
              <v-img :src="(skew === 1) ? '/skew0.jpg' : '/skew0-45.jpg'" contain max-height="450"></v-img>
            </v-col>
          </v-row>
        </v-container>
      </v-card-text>
    </v-card>
  </v-container>
</template>

<script>
import moment from 'moment-timezone'

export default {
  name: 'HomeView',
  data: () => ({
    toDay: moment().tz('Asia/Bangkok').locale('th').format('DD MMMM YYYY'),
    valid: false,
    pricePerUnits: [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
    ln: 0,
    w: 0,
    skew: 1,
    footingt38: [
      [14, 18],
      [16, 20]
    ],
    wExp: [14, 17],
    w2Exp: [13, 16],
    wFLExp: [14, 17],
    p1: [5.5, 5.23, 4.97, 4.71, 4.45, 4.19, 3.94, 3.68, 3.42, 3.16, 2.9, 2.64, 2.38, 2.4],
    r1: [24, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 2, 10],
    t: [17, 1, 1, 4, 17, 1, 5],
    u: [3.08, 2.7, 2.38, 2.35, 3.03, 2.65, 2.31]
  }),
  methods: {
    numberFormatter (v) {
      return new Intl.NumberFormat('en-US', { minimumFractionDigits: 3 }).format(v)
    },
    popupFooting () {
      this.$swal({
        imageUrl: '/footing.jpg'
      })
    },
    popupAbutment () {
      this.$swal({
        imageUrl: '/abutment.jpg'
      })
    },
    focusNextInput (index) {
      if (index < this.$refs.form.inputs.length - 1) {
        this.$refs.form.inputs[index + 1].focus()
      } else {
        this.$refs.form.inputs[0].focus()
      }
    }
  },
  computed: {
    calcDB20: function () {
      return (((Math.ceil(((3 * this.skew) - 0.1) / 0.17)) + 1) * (((12.53 + (this.wExp[this.w] - 11)) * this.skew + (0.02 * 40)) * 2.466)) + (((Math.ceil(((3 * this.skew) - 0.1) / 0.3)) + 1) * (((11.15 + (this.wExp[this.w] - 11)) * this.skew + (0.02 * 40)) * 2.466)) + (((Math.ceil((((this.wExp[this.w] * this.skew)) - 0.1) / 0.3)) + 1) * ((3.65 * this.skew) * 2.466))
    },
    calcDB25: function () {
      return ((((Math.ceil((((this.wExp[this.w] * this.skew)) - 0.1) / 0.25)) + 1)) * ((5.01 * this.skew) * 3.853))
    },
    calcDB12: function () {
      let t164 = 0
      for (let i = 0; i < this.p1.length; i++) {
        t164 += ((this.p1[i] * this.skew) * (this.r1[i]) * 0.888)
      }
      return (((((1.38 * this.skew) * 0.888)) * ((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)) + (215.9438) + (t164) + (15.20256))
    },
    calcDB16: function () {
      const h36 = ((((9.9 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 4)
      const h37 = ((((11.31 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 9)
      const h38 = ((((10.9 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 9)
      const h39 = ((((9.9 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 6)
      const h40 = ((((11.31 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 4)
      const h41 = ((((9.9 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 11)
      const h42 = ((((9.9 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 11)
      const h43 = ((((9.9 + (this.w2Exp[this.w] - 10)) * this.skew + (0.016 * 40)) * 1.578) * 11)
      const h45 = (((this.ln === 0 ? 2.5 : (2.5 - 0.45)) * 1.578) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)))
      const h47 = (((1.01 * this.skew) * 1.578) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)))
      const h51 = (((this.ln === 0 ? 3.76 : (3.76 - 0.45)) * 1.578) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)))
      let x209 = 0
      for (let i = 0; i < this.t.length; i++) {
        x209 += ((this.t[i]) * (this.u[i] * this.skew) * 1.578)
      }
      return (h36 + h37 + h38 + h39 + h40 + h41 + h42 + h43 + h45 + h47 + h51 + x209)
    },
    calcDB202: function () {
      const h44 = (((this.ln === 0 ? 2.5 : (2.5 - 0.45)) * 2.466) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)))
      const h48 = (((this.ln === 0 ? 2.43 : (2.43 + 0.45)) * 2.466) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)))
      const h49 = (((this.ln === 0 ? 2.07 : (2.07 + 0.45)) * 2.466) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.2)) + 1)))
      const h50 = (((this.ln === 0 ? 3.76 : (3.76 - 0.45)) * 2.466) * (((Math.ceil(((this.w2Exp[this.w] * this.skew) - 0.1) / 0.1)) + 1)))
      return (h44 + h48 + h49 + h50)
    },
    calcLeanConcrete: function () {
      const t33 = (0.05 + (3 * this.skew) + 0.05)
      const u33 = (0.05 + (this.wFLExp[this.w] * this.skew) + 0.05)
      return (t33 * u33 * 1)
    },
    calcCompactSand: function () {
      const t36 = (0.05 + (3 * this.skew) + 0.05)
      const u36 = (0.05 + (this.wFLExp[this.w] * this.skew) + 0.05)
      return (t36 * u36 * 1)
    },
    calcFormWork: function () {
      const formworks = [(this.wFLExp[this.w] * this.skew), (this.wFLExp[this.w] * this.skew), (3 * this.skew), (3 * this.skew)]
      let v30 = 0
      for (let i = 0; i < formworks.length; i++) {
        v30 += formworks[i]
      }
      return v30
    },
    calcConcrete35: function () {
      const s21 = (this.wFLExp[this.w] * this.skew)
      const t21 = 3
      const u21 = 1
      return (s21 * t21 * u21)
    },
    calcFormWork2: function () {
      const n36 = (((0.5 * (0.3 + 0.5) * (0.3 * this.skew))) + (0.5 * ((1.68 + 0.6 + 0.25) + 1.68 + 0.6) * 0.3 * this.skew) + (3.82 * (0.6 * this.skew))) * 2
      const n27 = (0.3 * (this.w2Exp[this.w] * this.skew))
      const n28 = (1.68 * (this.w2Exp[this.w] * this.skew))
      const n29 = (3.82 * (this.w2Exp[this.w] * this.skew))
      const l30 = (0.3 + (Math.sqrt(Math.pow((0.3 * this.skew), 2) + Math.pow(0.2, 2))) + (1.68 - 0.8) + 0.6 + (Math.sqrt(Math.pow((0.3 * this.skew), 2) + Math.pow(0.25, 2))) + (3.82 - 0.85))
      const n30 = (l30 * (this.w2Exp[this.w] * this.skew))
      const n37 = n36 + n27 + n28 + n29 + n30
      const v27 = ((39.666) + (0.3 * this.skew * 5.5) + (0.5 * this.skew * 5.5) + (0.25 * this.skew * 2.15) + ((Math.sqrt((Math.pow(0.25, 2)) + (Math.pow((0.25 * this.skew), 2)))) * 5.5) + (6.043 * 0.25) + (0.25 * this.skew * 1))
      return (n37 + v27)
    },
    calcConcrete352: function () {
      const skewNum = [{ text: '0', value: 1 }, { text: '5', value: 1.0038 }, { text: '10', value: 1.0154 }, { text: '15', value: 1.0353 }, { text: '20', value: 1.0642 }, { text: '25', value: 1.1034 }, { text: '30', value: 1.1547 }, { text: '35', value: 1.2208 }, { text: '40', value: 1.3054 }, { text: '45', value: 1.4142 }].find(v => {
        return v.value === this.skew
      })
      const n18 = (0.120 * (this.w2Exp[this.w] * this.skew))
      const o18 = (Math.tan(parseInt(skewNum.text) * (Math.PI / 180)) * 0.120)
      const p18 = ((n18) - (o18))
      const n19 = ((0.5 * ((1.68 + 0.6 + 0.25) + 1.68 + 0.6) * 0.3) * (this.w2Exp[this.w] * this.skew))
      const o19 = (Math.tan(parseInt(skewNum.text) * (Math.PI / 180)) * (0.5 * ((1.68 + 0.6 + 0.25) + 1.68 + 0.6) * 0.3))
      const p19 = n19 - o19
      const n20 = (3.82 * 0.6) * (this.w2Exp[this.w] * this.skew)
      const o20 = (Math.tan(parseInt(skewNum.text) * (Math.PI / 180)) * (3.82 * 0.6))
      const p20 = n20 - o20
      const p21 = p18 + p19 + p20
      const v16 = (((0.5 * this.skew) * 0.3 * 5.5) + (0.5 * 0.25 * (0.25 - (0.25 * this.skew - 0.25)) * 5.5) + ((5 * 0.3 * 5.5) - ((3.525 * 0.25 * 2.35) / 2) - (3.525 * 0.25))) * 2
      return p21 + v16
    }
  }
}
</script>
