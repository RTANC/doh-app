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
  },
  {
    path: '/SoilAggregateSubbase',
    name: 'SoilAggregateSubbaseView',
    component: () => import('@/views/SoilAggregateSubbaseView.vue')
  },
  {
    path: '/CrushedRockSoilAggregateTypeBase',
    name: 'CrushedRockSoilAggregateTypeBaseView',
    component: () => import('@/views/CrushedRockSoilAggregateTypeBaseView.vue')
  },
  {
    path: '/MillingOfExistingAsphaltSurface',
    name: 'MillingOfExistingAsphaltSurfaceView',
    component: () => import('@/views/MillingOfExistingAsphaltSurfaceView.vue')
  },
  {
    path: '/MillingOfExistingAsphaltSurface10',
    name: 'MillingOfExistingAsphaltSurface10View',
    component: () => import('@/views/MillingOfExistingAsphaltSurface10View.vue')
  }
]

const router = new VueRouter({
  mode: 'history',
  base: process.env.BASE_URL,
  routes
})

export default router
