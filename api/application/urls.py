from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('about', views.about, name='about'),
    path('services', views.services, name='services'),
    path('contact', views.contact, name='contact'),
    path('login', views.login, name='login'),
    path('singup', views.inscrir, name='singup'),
    path('main', views.main, name='main'),
    path('logout/', views.logout_view, name='logout'),




    path('vehicules', views.vehicules, name='vehicules'),
    path('vehicules/<int:vehicule_id>/', views.details_vehicule, name='details_vehicule'),
    path('VerifierDate/', views.VerifierDate, name='VerifierDate'),
    path('contrat_vehicule', views.contrat_vehicule, name='contrat_vehicule'),
    path('contrat_doc/<int:contrat_id>/', views.contrat_doc, name='contrat_doc'),
    path('confirmerReservation', views.confirmerReservation, name='confirmerReservation'),
    path('annulerReservation', views.annulerReservation, name='annulerReservation'),


]