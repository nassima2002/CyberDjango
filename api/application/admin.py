from django.contrib import admin

# Register your models here.
from django.contrib import admin
from .models import ContactForm, Vehicule, Contrat
from .models import Registre




@admin.register(ContactForm)
class RequestDemoAdmin(admin.ModelAdmin):
    list_display = [field.name for field in
                    ContactForm._meta.get_fields()]

@admin.register(Registre)
class RequestsDemoAdmin(admin.ModelAdmin):
        list_display = [field.name for field in
                        Registre._meta.get_fields()]


# Champs à afficher dans la liste des véhicules
class VehiculeAdmin(admin.ModelAdmin):
    list_display = ('id','nom', 'marque', 'prix_regulier','date_ajout')

admin.site.register(Vehicule, VehiculeAdmin)

admin.site.register(Contrat)
