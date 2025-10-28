from django.db import models
from django.contrib.auth.forms import PasswordResetForm
from django.urls import reverse


# Create your models here.

class ContactForm(models.Model):
    name= models.CharField(max_length=100)
    lastname = models.CharField(max_length=100)
    email= models.EmailField()
    message= models.CharField(max_length=100)
    subject= models.CharField(max_length=100)
    class Meta:
         db_table="client"

class Registre(models.Model):
    name = models.CharField(max_length=100)
    number = models.IntegerField()
    email = models.EmailField()
    password = models.CharField(max_length=100)
    class Meta:
        db_table = "inscription"


class Login(models.Model):
    name = models.CharField(max_length=100)
    password = models.CharField(max_length=100)
    class Meta:
        db_table = "login"






class Vehicule(models.Model):
    nom = models.CharField(max_length=100)
    marque = models.CharField(max_length=100)
    modele = models.CharField(max_length=100)
    annee = models.IntegerField()
    image = models.ImageField(upload_to='static/image_vehicule/')
    prix_regulier = models.DecimalField(max_digits=8, decimal_places=2)
    prix_promo = models.DecimalField(max_digits=8, decimal_places=2, null=True, blank=True)
    couleur = models.CharField(max_length=100)
    kilometrage = models.PositiveIntegerField()
    nombre_de_place = models.PositiveIntegerField()
    carburant = models.CharField(max_length=100)
    puissance = models.PositiveIntegerField()
    transmission = models.CharField(max_length=100)

    date_ajout = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.nom} {self.modele} ({self.annee})"

    def get_absolute_url(self):
        return reverse('details_vehicule', kwargs={'pk': self.pk})
    class Meta:
        db_table = "vehicule"



class Contrat(models.Model):
    vehicule = models.ForeignKey(Vehicule, on_delete=models.CASCADE)
    client = models.ForeignKey(Registre, on_delete=models.CASCADE)
    date_debut = models.DateField()
    date_fin = models.DateField()
    prix_total = models.DecimalField(max_digits=8, decimal_places=2)
    date_reservation = models.DateField(auto_now_add=True)

    def nom_client(self):
        return self.client.name

    def numero_client(self):
        return self.client.number
    def email_client(self):
        return self.client.email

    def __str__(self):
        return f"Contrat #{self.id} -- Nom Véhicule: {self.vehicule.nom} -- Client: {self.nom_client()} -- Email: {self.email_client()} -- Téléphone: {self.numero_client()}"
    class Meta:
        db_table = "contrat"
