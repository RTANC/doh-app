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
                <span>2 &nbsp;</span>
                <span class="text-decoration-underline">Soil Aggregate Subbase</span>
            </v-card-title>
            <v-card-text>
                <v-form ref="form">
                    <v-container fluid>
                        <v-row>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                                <p class="text-body-1 font-weight-medium black--text">ค่าวัสดุจากแหล่ง</p>
                            </v-col>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                                <div class="d-flex justify-space-between align-baseline">
                                  <v-icon>mdi-equal</v-icon>
                                  <span style="width: 80%;">
                                    <v-text-field type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)" autofocus></v-text-field>
                                  </span>
                                  <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                                </div>
                            </v-col>
                        </v-row>
                        <v-row>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                                <p class="text-body-1 font-weight-medium black--text">ค่าดำเนินการ + ค่าเสื่อมราคา (ขุด-ชน)</p>
                            </v-col>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                                <div class="d-flex justify-space-between align-baseline">
                                  <v-icon>mdi-equal</v-icon>
                                  <span style="width: 80%;">
                                    <v-text-field type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)"></v-text-field>
                                  </span>
                                  <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                                </div>
                            </v-col>
                        </v-row>
                        <v-row>
                            <v-col cols="12" xs="3" sm="3" md="3" lg="3">
                                <p class="text-body-1 font-weight-medium black--text">ค่าขนส่ง</p>
                            </v-col>
                            <v-col cols="12" xs="3" sm="3" md="3" lg="3">
                                <div class="d-flex justify-space-between align-baseline">
                                  <span style="width: 80%;">
                                    <v-text-field type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)"></v-text-field>
                                  </span>
                                  <span class="text-body-1 black--text font-weight-medium">กม.</span>
                                </div>
                            </v-col>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                              <div class="d-flex justify-space-between">
                                <v-icon>mdi-equal</v-icon>
                                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 1px solid gray; width: 70%;"></span>
                                <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                              </div>
                            </v-col>
                        </v-row>
                        <v-row>
                          <v-col cols="12" xs="1" sm="1" md="6" lg="6">
                            <p class="text-body-1 font-weight-medium black--text">รวม</p>
                          </v-col>
                          <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                            <div class="d-flex justify-space-between">
                              <v-icon>mdi-equal</v-icon>
                              <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 3px double gray; width: 70%;"></span>
                              <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                            </div>
                          </v-col>
                        </v-row>
                        <v-row>
                            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
                                <p class="text-body-1 font-weight-medium black--text">ส่วนยุบตัว</p>
                            </v-col>
                            <v-col cols="12" xs="2" sm="2" md="2" lg="2">
                                <v-text-field type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)"></v-text-field>
                            </v-col>
                            <v-col cols="12" xs="1" sm="1" md="2" lg="2">
                                <v-text-field type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)" prepend-icon="mdi-close"></v-text-field>
                            </v-col>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                              <div class="d-flex justify-space-between">
                                <v-icon>mdi-equal</v-icon>
                                <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 1px solid gray; width: 70%;"></span>
                                <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                              </div>
                            </v-col>
                        </v-row>
                        <v-row>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                                <p class="text-body-1 font-weight-medium black--text">ค่าดำเนินการ + ค่าเสื่อมราคา (บดทับ)</p>
                            </v-col>
                            <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                                <div class="d-flex justify-space-between align-baseline">
                                  <v-icon>mdi-equal</v-icon>
                                  <span style="width: 80%;">
                                    <v-text-field type="number" filled dense reverse hide-spin-buttons @keydown.enter="focusNextInput(0)"></v-text-field>
                                  </span>
                                  <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                                </div>
                            </v-col>
                        </v-row>
                        <v-row>
                          <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                            <p class="text-body-1 font-weight-medium black--text d-flex justify-end">ค่างานต้นทุนรวม</p>
                          </v-col>
                          <v-col cols="12" xs="6" sm="6" md="6" lg="6">
                            <div class="d-flex justify-space-between">
                              <v-icon>mdi-equal</v-icon>
                              <span class="text-body-1 black--text font-weight-medium d-flex justify-center" style="border-bottom: 3px double gray; width: 70%;"></span>
                              <span class="text-body-1 black--text font-weight-medium">บาท/ลบ.ม.</span>
                            </div>
                          </v-col>
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
  name: 'SoilAggregateSubbaseView',
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
      } catch (error) {
        console.log(error)
      }
    },
    exportToExcel () {
      // eslint-disable-next-line quotes
      const table = ``
      const wb = XLSX.read(table, { type: 'string' })
      XLSX.writeFile(wb, 'Soil Aggregate Subbase.xlsx', { raw: true })
    }
  }
}
</script>
