import Vue from 'vue'
import VueRouter from 'vue-router'

Vue.use(VueRouter)

const routes = [
  {
    path: '/',
    name: 'home',
    component: () => import('@/views/HomeView.vue')
  },
  {
    path: '/AsphaltConcreteLeveling',
    name: 'AsphaltConcreteLevelingView',
    component: () => import('@/views/AsphaltConcreteLevelingView.vue')
  },
  {
    path: '/SoftMaterialExcavation',
    name: 'SoftMaterialExcavationView',
    component: () => import('@/views/SoftMaterialExcavationView.vue')
  }
]

const router = new VueRouter({
  mode: 'history',
  base: process.env.BASE_URL,
  routes
})

export default router
