<template>
  <div id="app">
    <XlsxDown />
  </div>
</template>

<script>
import XlsxDown from './components/XlsxDown.vue'

export default {
  name: 'App',
  components: {
    XlsxDown
  }
}
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
