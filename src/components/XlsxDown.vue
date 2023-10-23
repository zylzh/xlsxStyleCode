<template>
  <div>
    <span class="btn" @click="createExcel">点击生成Excel</span>
  </div>
</template>

<script>
import { exportExcel } from '@/utils'
export default {
  name: 'XlsxDown',
  props: {
    msg: String
  },
  methods: {
    createExcel() {
      // 测试数据
      const testData = [
        {
          '序号': 1,
          '表头1': 1,
          '表头2': 2,
          '表头3': 3,
          '表头4': 4,
          '表头5': 5,
        },
        {
          '序号': 2,
          '表头1': 11,
          '表头2': 22,
          '表头3': 33,
          '表头4': 44,
          '表头5': 55,
        },
        {
          '序号': 3,
          '表头1': 111,
          '表头2': 222,
          '表头3': 333,
          '表头4': 444,
          '表头5': 555,
        },
      ]
      const temp = {
        '序号': testData.length + 1
      }
      temp['表头4'] = 0
      temp['表头5'] = 0
      testData.map((item) => {
        temp['表头4'] += parseFloat(item['表头4'])
        temp['表头5'] += parseFloat(item['表头5'])
      })
      testData.push(temp)
      exportExcel(testData, 'Excel下载', Object.keys(testData[0]).length - 1)
    }
  }
}
</script>

<!-- Add "scoped" attribute to limit CSS to this component only -->
<style scoped>
  .btn {
    border: 1px solid rgba(30,128,255,.3);
    background-color: rgba(30,128,255,.05);
    border-radius: 4px;
    cursor: pointer;
    padding: 10px;
    color: #1e80ff;
    font-size: 14px;
  }
</style>
