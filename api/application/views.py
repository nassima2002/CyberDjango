import os
from datetime import datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.core.paginator import Paginator
from django.db.models import Q
from django.contrib.auth import logout
from docx import Document
import pythoncom
import docx2pdf

from .models import ContactForm, Registre, Vehicule, Contrat


# ------------------------------
# PAGES PUBLIQUES
# ------------------------------
def home(request):
    return render(request, 'home.html')


def about(request):
    return render(request, 'about.html')


def services(request):
    return render(request, 'services.html')


# ------------------------------
# AUTHENTIFICATION UTILISATEUR
# ------------------------------
def inscrir(request):
    if request.method == "POST":
        name = request.POST['name']
        number = request.POST['number']
        email = request.POST['email']
        password = request.POST['password']

        if Registre.objects.filter(email=email).exists():
            return redirect('singup')

        utilisateur = Registre(name=name, number=number, email=email, password=password)
        utilisateur.save()
        return render(request, 'login.html')

    return render(request, 'singup.html')

# ------------------------------
# vulnurabiliter 
# ------------------------------

#def login(request):
    if request.method == "POST":
        username = request.POST.get("name")
        password = request.POST.get("password")

        # ⚠️ Vulnérabilité : injection SQL possible
        query = f"SELECT * FROM inscription WHERE name = '{username}' AND password = '{password}'"
        with connection.cursor() as cursor:
            cursor.execute(query)
            user = cursor.fetchone()

        if user:
            return JsonResponse({"status": "success", "message": "Connexion réussie"})
        else:
            return JsonResponse({"status": "error", "message": "Nom d'utilisateur ou mot de passe incorrect"})
    return JsonResponse({"error": "Méthode non autorisée"}, status=405)


def login(request):
    if request.method == 'POST':
        name = request.POST.get('name')
        password = request.POST.get('password')

        try:
            utilisateur = Registre.objects.get(name=name)
        except Registre.DoesNotExist:
            return redirect('login')

        if utilisateur.password == password:
            request.session['id'] = utilisateur.id
            return redirect('main')

        return redirect('login')

    return render(request, 'login.html')


def logout_view(request):
    logout(request)
    request.session.pop('id', None)
    return redirect('login')


# ------------------------------
# PAGE PRINCIPALE UTILISATEUR
# ------------------------------
def main(request):
    id = request.session.get('id')
    if not id:
        return redirect('login')

    utilisateur = Registre.objects.get(pk=id)
    contrat = Contrat.objects.filter(client=utilisateur).order_by('-date_reservation')

    return render(request, 'main.html', {'utilisateur': utilisateur, 'contrat': contrat})


# ------------------------------
# PAGE CONTACT
# ------------------------------
def contact(request):
    if request.method == "POST":
        name = request.POST['name']
        lastname = request.POST['lastname']
        email = request.POST['email']
        message = request.POST['message']
        subject = request.POST['subject']

        if not all([name, lastname, email, message, subject]):
            context = "Erreur, veuillez vérifier vos données."
        else:
            ContactForm.objects.create(
                name=name,
                lastname=lastname,
                email=email,
                message=message,
                subject=subject
            )
            context = "Les données ont été validées avec succès."

        return render(request, 'contact.html', {'context': context})

    return render(request, 'contact.html')


# ------------------------------
# VEHICULES
# ------------------------------
def vehicules(request):
    marque = request.GET.get('marque', '')
    type_vehicule = request.GET.get('type_vehicule', '')
    annee = request.GET.get('annee', '')

    annee = int(float(annee)) if annee else 0

    vehicules = Vehicule.objects.all().order_by('-date_ajout')

    if marque:
        vehicules = vehicules.filter(marque=marque)
    if annee:
        vehicules = vehicules.filter(annee=annee)
    if type_vehicule:
        vehicules = vehicules.filter(carburant=type_vehicule)

    paginator = Paginator(vehicules, 4)
    page_obj = paginator.get_page(request.GET.get('page'))

    context = {
        'marques': Vehicule.objects.values_list('marque', flat=True).distinct(),
        'annees': Vehicule.objects.values_list('annee', flat=True).distinct(),
        'types_vehicules': Vehicule.objects.values_list('carburant', flat=True).distinct(),
        'page_obj': page_obj,
        'vehicules': vehicules,
        'marque_search': marque,
        'annee_search': annee,
        'type_vehicule_search': type_vehicule,
    }
    return render(request, 'vehicule.html', context)


def details_vehicule(request, vehicule_id):
    vehicule = get_object_or_404(Vehicule, id=vehicule_id)
    return render(request, 'details_vehicule.html', {'vehicule': vehicule})


# ------------------------------
# FONCTIONS UTILITAIRES
# ------------------------------
def calcul_nbre_jours(date_debut, date_fin):
    date_debut = datetime.strptime(date_debut, "%Y-%m-%d")
    date_fin = datetime.strptime(date_fin, "%Y-%m-%d")
    return (date_fin - date_debut).days


def VerificationFunct(vehicule_id, date_debut, date_fin):
    vehicule = Vehicule.objects.get(id=vehicule_id)
    date_debut_client = datetime.strptime(date_debut, "%Y-%m-%d").date()
    date_fin_client = datetime.strptime(date_fin, "%Y-%m-%d").date()

    locations_existantes = Contrat.objects.filter(
        Q(date_debut__lte=date_fin_client, date_fin__gte=date_debut_client),
        vehicule=vehicule
    ).exists()

    return not locations_existantes


# ------------------------------
# RESERVATION
# ------------------------------
def VerifierDate(request):
    id = request.session.get('id')
    if not id:
        return JsonResponse({'disponible': False, 'id': None})

    utilisateur = Registre.objects.get(pk=id)
    vehicule_id = request.GET.get('vehicule_id')
    date_debut = request.GET.get('dateDeb')
    date_fin = request.GET.get('dateFin')

    disponible = VerificationFunct(vehicule_id, date_debut, date_fin)

    request.session['vehicule_id'] = vehicule_id
    request.session['date_debut'] = date_debut
    request.session['date_fin'] = date_fin

    return JsonResponse({'disponible': disponible, 'id': id})


def contrat_vehicule(request):
    vehicule_id = request.session.get('vehicule_id')
    id = request.session.get('id')
    date_debut = request.session.get('date_debut')
    date_fin = request.session.get('date_fin')

    utilisateur = Registre.objects.get(pk=id)
    vehicule = Vehicule.objects.get(id=vehicule_id)
    nbre_jours = calcul_nbre_jours(date_debut, date_fin)

    prix_vehicule = vehicule.prix_promo or vehicule.prix_regulier
    prix_total = prix_vehicule * nbre_jours

    context = {
        'vehicule': vehicule,
        'client': utilisateur,
        'date_debut': date_debut,
        'date_fin': date_fin,
        'prix_total': prix_total,
        'prix_vehicule': prix_vehicule,
    }
    return render(request, 'contrat_vehicule.html', context)


def confirmerReservation(request):
    vehicule_id = request.session.get('vehicule_id')
    id = request.session.get('id')
    date_debut = request.session.get('date_debut')
    date_fin = request.session.get('date_fin')

    utilisateur = Registre.objects.get(pk=id)
    vehicule = Vehicule.objects.get(id=vehicule_id)
    nbre_jours = calcul_nbre_jours(date_debut, date_fin)

    prix_total = (vehicule.prix_promo or vehicule.prix_regulier) * nbre_jours

    contrat = Contrat.objects.create(
        vehicule=vehicule,
        client=utilisateur,
        date_debut=date_debut,
        date_fin=date_fin,
        prix_total=prix_total
    )

    return render(request, 'contrat_doc.html', {'contrat': contrat})


def annulerReservation(request):
    for key in ['vehicule_id', 'date_debut', 'date_fin']:
        request.session.pop(key, None)
    return redirect('vehicules')


# ------------------------------
# CONTRAT - EXPORT DOCX/PDF
# ------------------------------
def contrat_doc(request, contrat_id):
    pythoncom.CoInitialize()
    contrat = get_object_or_404(Contrat, id=contrat_id)
    client = contrat.client
    vehicule = contrat.vehicule

    prix_vehicule = vehicule.prix_promo or vehicule.prix_regulier

    modele_word_path = os.path.join('static/Contrats/Contrat_Véhicule.docx')
    nom_fichier = f"contrat_{client.name}_{vehicule.nom}.docx"
    contrat_word_path = os.path.join('static/Contrats/', nom_fichier)

    modele_word = Document(modele_word_path)
    variables = {
        'nom_client': client.name,
        'email': client.email,
        'number': client.number,
        'nom_vehicule': vehicule.nom,
        'marque': vehicule.marque,
        'modele': vehicule.modele,
        'prix_vehicule': prix_vehicule,
        'date_debut': contrat.date_debut,
        'date_fin': contrat.date_fin,
        'date_reservation': contrat.date_reservation,
        'prix_total': str(contrat.prix_total),
        'num_contrat': contrat.id,
    }

    for paragraph in modele_word.paragraphs:
        for variable, valeur in variables.items():
            if variable in paragraph.text:
                paragraph.text = paragraph.text.replace(variable, str(valeur))

    modele_word.save(contrat_word_path)

    contrat_pdf_path = os.path.join('static/Contrats/', f"contrat_{client.name}_{vehicule.nom}.pdf")
    docx2pdf.convert(contrat_word_path, contrat_pdf_path)

    with open(contrat_pdf_path, 'rb') as f:
        response = HttpResponse(f, content_type='application/pdf')
        response['Content-Disposition'] = f'attachment; filename="contrat_{client.name}_{vehicule.nom}.pdf"'
        return response
