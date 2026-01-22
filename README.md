# ENSPD PV Management System

Système de gestion des Procès-Verbaux de délibération pour l'École Nationale Supérieure Polytechnique de Douala (ENSPD).

## 📋 Fonctionnalités

- ✅ **Import Excel** - Import de fichiers Excel PV avec parser intelligent
- ✅ **Dashboard interactif** - Filtres dynamiques (UE, ECUE, décision, recherche, moyenne)
- ✅ **Export Excel complet** - Export avec toutes les notes détaillées par UE et ECUE
- ✅ **Émargement avec filtre** - Feuilles d'émargement avec filtres appliqués et matière
- ✅ **Émargements NV** - Export multi-feuilles des étudiants Non Validés par matière
- ✅ **Émargements V et VC** - Export multi-feuilles des étudiants Validés par matière
- ✅ **Vue impression** - Vue optimisée pour l'impression des PV
- ✅ **Interface responsive** - Design moderne avec Tailwind CSS

## 🚀 Technologies

- **Backend**: Django 5.2+
- **Python**: 3.10+
- **Excel**: openpyxl pour la manipulation de fichiers Excel
- **Frontend**: Tailwind CSS, jQuery
- **Base de données**: SQLite (dev) / PostgreSQL (production)

## 📦 Installation

### Prérequis

- Python 3.10 ou supérieur
- pip (gestionnaire de paquets Python)

### Étapes d'installation

1. **Cloner le repository**

```bash
git clone https://github.com/VOTRE_USERNAME/enspd-pv-management.git
cd enspd-pv-management
```

2. **Créer un environnement virtuel**

```bash
python -m venv venv
```

3. **Activer l'environnement virtuel**

- Windows:
```bash
venv\Scripts\activate
```

- Linux/Mac:
```bash
source venv/bin/activate
```

4. **Installer les dépendances**

```bash
pip install -r requirements.txt
```

5. **Appliquer les migrations**

```bash
python manage.py migrate
```

6. **Créer un superutilisateur (optionnel)**

```bash
python manage.py createsuperuser
```

7. **Lancer le serveur de développement**

```bash
python manage.py runserver
```

8. **Accéder à l'application**

Ouvrez votre navigateur et accédez à : `http://127.0.0.1:8000/`

## 📖 Utilisation

### Import d'un PV

1. Accédez à la page d'accueil
2. Cliquez sur "Importer un PV"
3. Sélectionnez votre fichier Excel (.xlsx)
4. Le système analysera et importera automatiquement les données

### Dashboard et filtres

- **Filtre par statut global** : Validés, Non Validés, Compensation
- **Filtre par UE** : Sélectionnez une Unité d'Enseignement
- **Filtre par ECUE** : Sélectionnez une matière spécifique
- **Filtre par statut matière** : V, NV ou VC dans une matière
- **Recherche** : Par nom ou matricule d'étudiant
- **Filtre par moyenne** : Min et Max

### Exports disponibles

1. **Exporter Excel** - Export complet avec toutes les colonnes
2. **Émargement avec filtre** - Feuille simple avec filtres appliqués
3. **Émargements NV** - Multi-feuilles des étudiants à rattraper
4. **Émargements V et VC** - Multi-feuilles des étudiants validés
5. **Imprimer** - Vue optimisée pour l'impression

## 🗂️ Structure du projet

```
ENSPD/
├── pv/                     # Application principale
│   ├── models.py           # Modèles de données
│   ├── views.py            # Vues et logique métier
│   ├── urls.py             # Routes URL
│   ├── forms.py            # Formulaires
│   ├── utils/              # Utilitaires
│   │   └── excel_parser.py # Parser Excel
│   └── templates/          # Templates HTML
│       └── pv/
│           ├── base.html
│           ├── home.html
│           ├── dashboard.html
│           └── print.html
├── pv_management/          # Configuration Django
│   ├── settings.py
│   ├── urls.py
│   └── wsgi.py
├── static/                 # Fichiers statiques
├── media/                  # Fichiers uploadés
├── requirements.txt        # Dépendances Python
└── manage.py              # Script de gestion Django
```

## 🔧 Configuration

### Variables d'environnement

Créez un fichier `.env` à la racine du projet:

```env
SECRET_KEY=votre-clé-secrète-django
DEBUG=True
ALLOWED_HOSTS=localhost,127.0.0.1
```

## 📝 Modèles de données

- **ProcesVerbal** : PV avec métadonnées (filière, niveau, semestre, année)
- **UE** : Unité d'Enseignement
- **ECUE** : Élément Constitutif d'UE (matière)
- **Etudiant** : Étudiant avec notes et décision
- **Note** : Note d'un étudiant dans une matière
- **SyntheseUE** : Synthèse d'un étudiant pour une UE

## 🤝 Contribution

Les contributions sont les bienvenues! Pour contribuer:

1. Fork le projet
2. Créez une branche pour votre fonctionnalité (`git checkout -b feature/nouvelle-fonctionnalite`)
3. Committez vos changements (`git commit -m 'Ajout nouvelle fonctionnalité'`)
4. Poussez vers la branche (`git push origin feature/nouvelle-fonctionnalite`)
5. Ouvrez une Pull Request

## 📄 Licence

Ce projet est développé pour l'ENSPD (École Nationale Supérieure Polytechnique de Douala).

## 👥 Auteur

**RYDI Group** - 2026

## 📞 Support

Pour toute question ou problème, veuillez créer une issue sur GitHub.
