<template>
  <v-app>
    <v-app-bar
      app
      color="blue darken-4"
      dark
      clipped-right
    >
      <div class="d-flex align-center headline text-uppercase font-weight-bold">
        <v-img
          alt="doh Logo"
          class="shrink mr-2"
          contain
          src="https://www.doh.go.th/template/template4_1/blue/images/logodoh_2565.png"
          transition="scale-transition"
          width="50"
        />
        <span class="hidden-sm-and-down">Department of highway</span>
        <span class="d-md-none">DOH</span>
      </div>

      <v-spacer></v-spacer>

      <v-app-bar-nav-icon @click.stop="drawer = !drawer"></v-app-bar-nav-icon>
    </v-app-bar>
    <v-navigation-drawer v-model="drawer" app clipped right dark color="blue darken-4">
      <v-list nav>
        <v-list-item-group color="blue lighten-5">
          <v-list-item :to="{ path: '/' }" @click.stop="drawer = false">
            <v-list-item-content>
              <v-list-item-title>หน้าแรก</v-list-item-title>
            </v-list-item-content>
          </v-list-item>
          <v-list-item :to="{ path: '/about' }" @click.stop="drawer = false">
            <v-list-item-content>
              <v-list-item-title>เกี่ยวกับ</v-list-item-title>
            </v-list-item-content>
          </v-list-item>
        </v-list-item-group>
      </v-list>
    </v-navigation-drawer>
    <v-main>
      <router-view/>
    </v-main>
  </v-app>
</template>

<script>

export default {
  name: 'App',

  data: () => ({
    drawer: false
  })
}
</script>
