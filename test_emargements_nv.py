"""
Script de test pour vérifier l'export des émargements NV
"""
import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'pv_management.settings')
django.setup()

from pv.models import ProcesVerbal, Note, ECUE, UE

print("="*80)
print("TEST DE L'EXPORT ÉMARGEMENTS NV")
print("="*80)

# Récupérer le premier PV
pv = ProcesVerbal.objects.first()

if not pv:
    print("\n❌ Aucun PV trouvé dans la base de données")
    print("Veuillez d'abord importer un fichier Excel PV")
    exit(1)

print(f"\n1. PV SÉLECTIONNÉ")
print("-"*80)
print(f"Filière: {pv.filiere}")
print(f"Niveau: {pv.niveau}")
print(f"Semestre: {pv.semestre}")
print(f"Année: {pv.annee_academique}")

print(f"\n2. ANALYSE DES NOTES NV PAR MATIÈRE")
print("-"*80)

total_matieres = 0
matieres_avec_nv = 0
total_etudiants_nv = 0

# Parcourir toutes les UE
for ue in pv.ues.all().order_by('ordre'):
    print(f"\nUE: {ue.code} - {ue.intitule}")

    # Parcourir tous les ECUE de l'UE
    for ecue in ue.ecues.all().order_by('ordre'):
        total_matieres += 1

        # Compter les étudiants NV pour cet ECUE
        notes_nv = Note.objects.filter(
            ecue=ecue,
            decision='NV'
        ).select_related('etudiant')

        nb_nv = notes_nv.count()

        if nb_nv > 0:
            matieres_avec_nv += 1
            total_etudiants_nv += nb_nv
            print(f"  + {ecue.code}: {nb_nv} etudiant(s) NV")

            # Afficher les 3 premiers étudiants
            for idx, note in enumerate(notes_nv[:3], 1):
                etudiant = note.etudiant
                print(f"     {idx}. {etudiant.nom_prenom} - CC:{note.cc}, EX:{note.examen}, MOY:{note.moyenne}")

            if nb_nv > 3:
                print(f"     ... et {nb_nv - 3} autre(s)")
        else:
            print(f"  - {ecue.code}: 0 etudiant NV (tous valides)")

print(f"\n3. STATISTIQUES GLOBALES")
print("-"*80)
print(f"Total de matières (ECUE): {total_matieres}")
print(f"Matières avec au moins 1 NV: {matieres_avec_nv}")
print(f"Matières sans NV: {total_matieres - matieres_avec_nv}")
print(f"Total d'étudiants NV (toutes matières): {total_etudiants_nv}")

print(f"\n4. PRÉVISION DU FICHIER EXCEL")
print("-"*80)
if matieres_avec_nv > 0:
    print(f"OK Le fichier Excel contiendra {matieres_avec_nv} feuille(s)")
    print(f"OK Chaque feuille listera les etudiants NV de la matiere")
    print(f"OK Structure: N, Matricule, Nom, CC, EX, Moyenne, Decision, Signature")
else:
    print("ATTENTION AUCUNE feuille ne sera generee (aucun etudiant NV)")
    print("Le fichier contiendra une feuille 'Information' indiquant l'absence de NV")

print(f"\n5. TEST DE LA VUE")
print("-"*80)
print(f"URL de test: http://127.0.0.1:5000/export-emargements-nv/{pv.pk}/")
print(f"Nom du fichier attendu: Emargements_NV_[filiere]_{pv.niveau}_{pv.semestre}_[date].xlsx")

print("\nOK Analyse terminee!")
print("\nPour tester:")
print(f"1. Accédez au dashboard: http://127.0.0.1:5000/dashboard/{pv.pk}/")
print("2. Cliquez sur le bouton 'Émargements NV' (rouge)")
print("3. Vérifiez le téléchargement du fichier Excel")
print("4. Ouvrez le fichier et vérifiez les feuilles générées")
