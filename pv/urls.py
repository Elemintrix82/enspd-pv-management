from django.urls import path
from . import views

app_name = 'pv'

urlpatterns = [
    path('', views.home, name='home'),
    path('import/', views.import_pv, name='import'),
    path('dashboard/<int:pk>/', views.dashboard, name='dashboard'),
    path('dashboard-aggrid/<int:pk>/', views.dashboard_aggrid, name='dashboard_aggrid'),
    path('export/<int:pk>/', views.export_excel, name='export'),
    path('export-emargement/<int:pk>/', views.export_feuille_emargement, name='export_emargement'),
    path('export-emargements-nv/<int:pk>/', views.export_emargements_nv_complets, name='export_emargements_nv'),
    path('export-emargements-v-vc/<int:pk>/', views.export_emargements_v_vc, name='export_emargements_v_vc'),
    path('print/<int:pk>/', views.print_view, name='print'),
]
