# views.py (complete, secure)
import os
import logging
from datetime import datetime
from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse, JsonResponse
from django.core.paginator import Paginator
from django.db.models import Q
from django.contrib.auth import logout
from django.contrib.auth.hashers import make_password, check_password
from django.views.decorators.csrf import csrf_protect
from django.core.validators import validate_email
from django.core.exceptions import ValidationError
from django.utils.text import slugify
from django.conf import settings
import re

# Optional Windows-only imports wrapped to avoid crash on non-windows
try:
    import pythoncom
except Exception:
    pythoncom = None
try:
    from docx import Document
    import docx2pdf
except Exception:
    Document = None
    docx2pdf = None

# Models (adapter si noms différents)
from .models import ContactForm, Registre, Vehicule, Contrat

# Logger configuration: récupère le logger nommé dans settings.LOGGING
logger = logging.getLogger('application')

# ------------------------------
# Helpers / utilitaires
# ------------------------------
def _is_valid_phone(number: str) -> bool:
    """
    Vérifie un numéro simple (ex: 061234567) — adapte le regex selon ton format.
    On accepte uniquement chiffres et éventuellement +.
    """
    number = number.strip()
    return bool(re.fullmatch(r'(\+?\d{8,15})', number))

def safe_join_static(*parts):
    """
    Construit un chemin sûr relatif au dossier static du projet.
    Utilise settings.BASE_DIR si disponible.
    """
    base = getattr(settings, 'BASE_DIR', None)
    if base:
        return os.path.join(base, *parts)
    return os.path.join(*parts)

def _ensure_logs_dir():
    # utile si tu veux logger vers des fichiers depuis views (mais normalement configuré dans settings)
    logs_dir = os.path.join(getattr(settings, 'BASE_DIR', '.'), 'logs')
    os.makedirs(logs_dir, exist_ok=True)
    return logs_dir

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
# AUTHENTIFICATION UTILISATEUR (SECURE)
# ------------------------------
@csrf_protect
def inscrir(request):
    """
    Endpoint d'inscription sécurisé : stocke le mot de passe haché.
    """
    if request.method == "POST":
        name = request.POST.get('name', '').strip()
        number = request.POST.get('number', '').strip()
        email = request.POST.get('email', '').strip()
        password = request.POST.get('password', '')

        logger.debug("Tentative d'inscription: name=%s, email=%s", name, email)

        # Validation minimale
        if not (name and number and email and password):
            logger.warning("Inscription échouée: champs manquants")
            return render(request, 'singup.html', {'context': 'Tous les champs sont requis.'})

        # Validate email
        try:
            validate_email(email)
        except ValidationError:
            logger.warning("Inscription échouée: email invalide %s", email)
            return render(request, 'singup.html', {'context': 'Email invalide.'})

        # Validate phone
        if not _is_valid_phone(number):
            logger.warning("Inscription échouée: numéro invalide %s", number)
            return render(request, 'singup.html', {'context': 'Numéro de téléphone invalide.'})

        # Vérifier doublon par email ou nom (selon politique)
        if Registre.objects.filter(email=email).exists():
            logger.info("Inscription échouée: email déjà utilisé %s", email)
            return render(request, 'singup.html', {'context': "Email déjà utilisé."})
        if Registre.objects.filter(name=name).exists():
            logger.info("Inscription échouée: nom déjà utilisé %s", name)
            return render(request, 'singup.html', {'context': "Nom déjà utilisé."})

        # Hacher le mot de passe avant de le stocker
        hashed = make_password(password)

        try:
            utilisateur = Registre(name=name, number=number, email=email, password=hashed)
            utilisateur.save()
            logger.info("Nouvel utilisateur créé: %s (id=%s)", name, utilisateur.id)
            return redirect('login')
        except Exception as e:
            logger.exception("Erreur lors de la création d'utilisateur: %s", e)
            return render(request, 'singup.html', {'context': "Erreur interne, réessayer plus tard."})

    return render(request, 'singup.html')


@csrf_protect
def login(request):
    """
    Endpoint de login sécurisé : utilise ORM + check_password.
    - Pour les requêtes normales (form submit) : redirect vers 'main' si succès.
    - Pour les requêtes API/AJAX : retourne JSON.
    """
    if request.method == "POST":
        username = request.POST.get("name", "").strip()
        password = request.POST.get("password", "")

        logger.debug("Tentative de login: username=%s", username)

        # Protection basique : empêcher des inputs vides
        if not username or not password:
            logger.warning("Login échoué: champs manquants pour username=%s", username)
            if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
                return JsonResponse({"status": "error", "message": "Champs manquants"}, status=400)
            return render(request, 'login.html', {'context': 'Champs manquants'})

        # Récupérer l'utilisateur via ORM (paramétré par Django)
        utilisateur = Registre.objects.filter(name=username).first()

        # vérifier que l'utilisateur existe et que le hash correspond
        if utilisateur and check_password(password, utilisateur.password):
            # authentification réussie : créer session
            request.session['id'] = utilisateur.id
            logger.info("Connexion réussie: %s (id=%s)", username, utilisateur.id)

            # réponse API
            if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
                return JsonResponse({"status": "success", "message": "Connexion réussie"})

            # redirection pour formulaire normal
            return redirect('main')
        else:
            logger.warning("Échec d'authentification pour: %s", username)
            if request.headers.get('x-requested-with') == 'XMLHttpRequest' or 'application/json' in request.headers.get('Accept', ''):
                return JsonResponse({"status": "error", "message": "Nom d'utilisateur ou mot de passe incorrect"}, status=401)
            return render(request, 'login.html', {'context': 'Nom d\'utilisateur ou mot de passe incorrect'})

    # GET -> afficher la page de connexion
    return render(request, 'login.html')


def logout_view(request):
    user_id = request.session.get('id')
    request.session.pop('id', None)
    logout(request)
    logger.info("Utilisateur déconnecté: id=%s", user_id)
    return redirect('login')


# ------------------------------
# PAGE PRINCIPALE UTILISATEUR
# ------------------------------
def main(request):
    id = request.session.get('id')
    if not id:
        logger.debug("Accès à main sans session, redirection vers login")
        return redirect('login')

    utilisateur = get_object_or_404(Registre, pk=id)
    contrat = Contrat.objects.filter(client=utilisateur).order_by('-date_reservation')
    return render(request, 'main.html', {'utilisateur': utilisateur, 'contrat': contrat})


# ------------------------------
# CONTACT
# ------------------------------
@csrf_protect
def contact(request):
    if request.method == "POST":
        name = request.POST.get('name', '').strip()
        lastname = request.POST.get('lastname', '').strip()
        email = request.POST.get('email', '').strip()
        message = request.POST.get('message', '').strip()
        subject = request.POST.get('subject', '').strip()

        logger.debug("Contact form submit: %s %s %s", name, lastname, email)

        if not all([name, lastname, email, message, subject]):
            logger.warning("Contact form invalid: champs manquants")
            context = "Erreur, veuillez vérifier vos données."
            return render(request, 'contact.html', {'context': context})

        try:
            validate_email(email)
        except ValidationError:
            logger.warning("Contact form: email invalide %s", email)
            return render(request, 'contact.html', {'context': 'Email invalide.'})

        try:
            ContactForm.objects.create(
                name=name,
                lastname=lastname,
                email=email,
                message=message,
                subject=subject
            )
            logger.info("Contact form saved: %s %s", name, lastname)
            context = "Les données ont été validées avec succès."
            return render(request, 'contact.html', {'context': context})
        except Exception as e:
            logger.exception("Erreur en sauvegardant contact form: %s", e)
            return render(request, 'contact.html', {'context': "Erreur interne, réessayer."})

    return render(request, 'contact.html')


# ------------------------------
# VEHICULES
# ------------------------------
def vehicules(request):
    marque = request.GET.get('marque', '').strip()
    type_vehicule = request.GET.get('type_vehicule', '').strip()
    annee = request.GET.get('annee', '').strip()

    try:
        annee_int = int(float(annee)) if annee else None
    except (ValueError, TypeError):
        annee_int = None

    vehicules_qs = Vehicule.objects.all().order_by('-date_ajout')

    if marque:
        vehicules_qs = vehicules_qs.filter(marque=marque)
    if annee_int:
        vehicules_qs = vehicules_qs.filter(annee=annee_int)
    if type_vehicule:
        vehicules_qs = vehicules_qs.filter(carburant=type_vehicule)

    paginator = Paginator(vehicules_qs, 4)
    page_obj = paginator.get_page(request.GET.get('page'))

    context = {
        'marques': Vehicule.objects.values_list('marque', flat=True).distinct(),
        'annees': Vehicule.objects.values_list('annee', flat=True).distinct(),
        'types_vehicules': Vehicule.objects.values_list('carburant', flat=True).distinct(),
        'page_obj': page_obj,
        'vehicules': vehicules_qs,
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
    """
    Vérifie si le véhicule est disponible entre date_debut et date_fin.
    Utilise correctement les opérateurs d'ORM (date_debut__lte, date_fin__gte).
    """
    vehicule = get_object_or_404(Vehicule, id=vehicule_id)
    date_debut_client = datetime.strptime(date_debut, "%Y-%m-%d").date()
    date_fin_client = datetime.strptime(date_fin, "%Y-%m-%d").date()

    # Vérifie chevauchement : si un contrat existe qui chevauche la période demandée
    locations_existantes = Contrat.objects.filter(
        Q(date_debut__lte=date_fin_client, date_fin__gte=date_debut_client),
        vehicule=vehicule
    ).exists()

    logger.debug(
        "VerificationFunct: vehicule_id=%s, date_debut=%s, date_fin=%s, existe=%s",
        vehicule_id, date_debut_client, date_fin_client, locations_existantes
    )
    return not locations_existantes


# ------------------------------
# RESERVATION
# ------------------------------
def VerifierDate(request):
    id = request.session.get('id')
    if not id:
        logger.debug("VerifierDate: utilisateur non connecté")
        return JsonResponse({'disponible': False, 'id': None})

    utilisateur = get_object_or_404(Registre, pk=id)
    vehicule_id = request.GET.get('vehicule_id')
    date_debut = request.GET.get('dateDeb')
    date_fin = request.GET.get('dateFin')

    if not all([vehicule_id, date_debut, date_fin]):
        logger.warning("VerifierDate: paramètres manquants")
        return JsonResponse({'disponible': False, 'id': id})

    try:
        disponible = VerificationFunct(vehicule_id, date_debut, date_fin)
    except Exception as e:
        logger.exception("Erreur VérifierDate: %s", e)
        return JsonResponse({'disponible': False, 'id': id})

    # stocker dans session pour utilisation lors de la réservation
    request.session['vehicule_id'] = vehicule_id
    request.session['date_debut'] = date_debut
    request.session['date_fin'] = date_fin

    return JsonResponse({'disponible': disponible, 'id': id})


def contrat_vehicule(request):
    vehicule_id = request.session.get('vehicule_id')
    id = request.session.get('id')
    date_debut = request.session.get('date_debut')
    date_fin = request.session.get('date_fin')

    if not all([vehicule_id, id, date_debut, date_fin]):
        logger.debug("contrat_vehicule: paramètres session manquants, redirection vers vehicules")
        return redirect('vehicules')

    utilisateur = get_object_or_404(Registre, pk=id)
    vehicule = get_object_or_404(Vehicule, id=vehicule_id)
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


@csrf_protect
def confirmerReservation(request):
    vehicule_id = request.session.get('vehicule_id')
    id = request.session.get('id')
    date_debut = request.session.get('date_debut')
    date_fin = request.session.get('date_fin')

    if not all([vehicule_id, id, date_debut, date_fin]):
        logger.warning("confirmerReservation: paramètres manquants")
        return redirect('vehicules')

    utilisateur = get_object_or_404(Registre, pk=id)
    vehicule = get_object_or_404(Vehicule, id=vehicule_id)
    nbre_jours = calcul_nbre_jours(date_debut, date_fin)

    prix_total = (vehicule.prix_promo or vehicule.prix_regulier) * nbre_jours

    try:
        contrat = Contrat.objects.create(
            vehicule=vehicule,
            client=utilisateur,
            date_debut=date_debut,
            date_fin=date_fin,
            prix_total=prix_total
        )
        logger.info("Contrat créé: id=%s, client=%s, vehicule=%s", contrat.id, utilisateur.id, vehicule.id)
        return render(request, 'contrat_doc.html', {'contrat': contrat})
    except Exception as e:
        logger.exception("Erreur lors de la création du contrat: %s", e)
        return redirect('vehicules')


def annulerReservation(request):
    for key in ['vehicule_id', 'date_debut', 'date_fin']:
        request.session.pop(key, None)
    logger.debug("annulerReservation: session cleared")
    return redirect('vehicules')


# ------------------------------
# CONTRAT - EXPORT DOCX/PDF
# ------------------------------
@csrf_protect
def contrat_doc(request, contrat_id):
    # Init COM seulement si pythoncom est dispo (Windows)
    if pythoncom:
        try:
            pythoncom.CoInitialize()
        except Exception:
            logger.debug("pythoncom.CoInitialize not needed or failed")

    contrat = get_object_or_404(Contrat, id=contrat_id)
    client = contrat.client
    vehicule = contrat.vehicule

    prix_vehicule = vehicule.prix_promo or vehicule.prix_regulier

    # Construire chemins via settings.BASE_DIR pour être sûr
    modele_word_path = safe_join_static('static', 'Contrats', 'Contrat_Véhicule.docx')
    # sanitize filename using slugify (évite caractères problématiques)
    client_slug = slugify(client.name) or str(client.id)
    veh_slug = slugify(getattr(vehicule, 'nom', str(vehicule.id)))
    nom_fichier = f"contrat_{client_slug}_{veh_slug}.docx"
    contrat_word_path = safe_join_static('static', 'Contrats', nom_fichier)

    # Assure-toi que le fichier template existe
    if not Document:
        logger.error("python-docx n'est pas installé ou indisponible")
        return HttpResponse("Fonctionnalité DOCX non disponible sur ce serveur.", status=500)

    if not os.path.exists(modele_word_path):
        logger.error("Modèle Word introuvable: %s", modele_word_path)
        return HttpResponse("Modèle Word introuvable.", status=500)

    try:
        modele_word = Document(modele_word_path)
    except Exception as e:
        logger.exception("Erreur ouverture template Word: %s", e)
        return HttpResponse("Erreur lors de la lecture du modèle Word.", status=500)

    variables = {
        'nom_client': client.name,
        'email': client.email,
        'number': client.number,
        'nom_vehicule': getattr(vehicule, 'nom', ''),
        'marque': getattr(vehicule, 'marque', ''),
        'modele': getattr(vehicule, 'modele', ''),
        'prix_vehicule': prix_vehicule,
        'date_debut': contrat.date_debut,
        'date_fin': contrat.date_fin,
        'date_reservation': contrat.date_reservation,
        'prix_total': str(contrat.prix_total),
        'num_contrat': contrat.id,
    }

    # Replace variables in the Word template (simple replace)
    try:
        for paragraph in modele_word.paragraphs:
            for variable, valeur in variables.items():
                if variable in paragraph.text:
                    paragraph.text = paragraph.text.replace(variable, str(valeur))

        modele_word.save(contrat_word_path)
        logger.info("Contrat DOCX créé: %s", contrat_word_path)
    except Exception as e:
        logger.exception("Erreur lors du remplissage du template Word: %s", e)
        return HttpResponse("Erreur lors de la génération du contrat.", status=500)

    contrat_pdf_path = safe_join_static('static', 'Contrats', f"contrat_{client_slug}_{veh_slug}.pdf")
    try:
        if docx2pdf:
            docx2pdf.convert(contrat_word_path, contrat_pdf_path)
            logger.info("Conversion DOCX->PDF réussie: %s", contrat_pdf_path)
        else:
            logger.warning("docx2pdf non disponible, pas de conversion en PDF.")
            # Optionnel: renvoyer le docx si on ne peut pas convertir
            with open(contrat_word_path, 'rb') as f:
                response = HttpResponse(f.read(), content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
                response['Content-Disposition'] = f'attachment; filename="{os.path.basename(contrat_word_path)}"'
                return response
    except Exception as e:
        logger.exception("Erreur lors de la conversion en PDF: %s", e)
        return HttpResponse(f"Erreur lors de la conversion en PDF: {e}", status=500)

    try:
        with open(contrat_pdf_path, 'rb') as f:
            response = HttpResponse(f.read(), content_type='application/pdf')
            response['Content-Disposition'] = f'attachment; filename="contrat_{client_slug}_{veh_slug}.pdf"'
            return response
    except FileNotFoundError:
        logger.exception("Fichier PDF introuvable après conversion: %s", contrat_pdf_path)
        return HttpResponse("Fichier PDF introuvable après conversion.", status=500)
    except Exception as e:
        logger.exception("Erreur lors de l'envoi du PDF: %s", e)
        return HttpResponse("Erreur lors de la lecture du PDF.", status=500)
