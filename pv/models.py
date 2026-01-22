from django.db import models
from django.core.validators import MinValueValidator, MaxValueValidator
from django.db.models import Sum


class ProcesVerbal(models.Model):
    """
    Modèle représentant un Procès-Verbal de délibération importé
    """
    fichier = models.FileField(upload_to='pv/', verbose_name="Fichier PV")
    filiere = models.CharField(max_length=200, verbose_name="Filière")
    niveau = models.IntegerField(verbose_name="Niveau d'étude")
    semestre = models.CharField(max_length=10, verbose_name="Semestre")
    annee_academique = models.CharField(max_length=20, verbose_name="Année académique")
    formation = models.CharField(max_length=100, verbose_name="Type de formation", blank=True, null=True)
    date_import = models.DateTimeField(auto_now_add=True, verbose_name="Date d'import")

    class Meta:
        verbose_name = "Procès-Verbal"
        verbose_name_plural = "Procès-Verbaux"
        ordering = ['-date_import']

    def __str__(self):
        return f"PV {self.filiere} - {self.niveau} - {self.semestre} ({self.annee_academique})"

    @property
    def nombre_etudiants(self):
        return self.etudiants.count()

    @property
    def nombre_valides(self):
        return self.etudiants.filter(decision_generale='V').count()

    @property
    def nombre_non_valides(self):
        return self.etudiants.filter(decision_generale='NV').count()

    @property
    def nombre_valides_compensation(self):
        return self.etudiants.filter(decision_generale='VC').count()

    @property
    def taux_reussite(self):
        total = self.nombre_etudiants
        if total == 0:
            return 0
        return round((self.nombre_valides + self.nombre_valides_compensation) / total * 100, 2)


class UE(models.Model):
    """
    Modèle représentant une Unité d'Enseignement (UE)
    """
    pv = models.ForeignKey(ProcesVerbal, on_delete=models.CASCADE, related_name='ues')
    code = models.CharField(max_length=50, verbose_name="Code UE")
    intitule = models.CharField(max_length=300, verbose_name="Intitulé")
    ordre = models.IntegerField(verbose_name="Ordre d'affichage")

    class Meta:
        verbose_name = "Unité d'Enseignement"
        verbose_name_plural = "Unités d'Enseignement"
        ordering = ['ordre']
        unique_together = ['pv', 'code']

    def __str__(self):
        return f"{self.code} - {self.intitule}"


class ECUE(models.Model):
    """
    Modèle représentant un Élément Constitutif d'UE (matière)
    """
    ue = models.ForeignKey(UE, on_delete=models.CASCADE, related_name='ecues')
    code = models.CharField(max_length=50, verbose_name="Code ECUE")
    intitule = models.CharField(max_length=300, verbose_name="Intitulé")
    ordre = models.IntegerField(verbose_name="Ordre d'affichage")
    credits = models.IntegerField(default=3, verbose_name="Nombre de crédits")

    class Meta:
        verbose_name = "Élément Constitutif d'UE"
        verbose_name_plural = "Éléments Constitutifs d'UE"
        ordering = ['ordre']
        unique_together = ['ue', 'code']

    def __str__(self):
        return f"{self.code} - {self.intitule}"


class Etudiant(models.Model):
    """
    Modèle représentant un étudiant avec ses informations et résultats globaux
    """
    DECISION_CHOICES = [
        ('V', 'Validé'),
        ('NV', 'Non Validé'),
        ('VC', 'Validé par Compensation'),
    ]

    pv = models.ForeignKey(ProcesVerbal, on_delete=models.CASCADE, related_name='etudiants')
    numero = models.IntegerField(verbose_name="N°")
    matricule = models.CharField(max_length=50, verbose_name="Matricule")
    nom_prenom = models.CharField(max_length=300, verbose_name="Nom & Prénom")
    moyenne_generale = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        validators=[MinValueValidator(0), MaxValueValidator(20)],
        verbose_name="Moyenne générale",
        null=True,
        blank=True,
        default=None
    )
    credits_acquis = models.IntegerField(
        default=0,
        verbose_name="Crédits acquis",
        null=True,
        blank=True
    )
    decision_generale = models.CharField(
        max_length=10,
        choices=DECISION_CHOICES,
        verbose_name="Décision générale",
        null=True,
        blank=True,
        default=None
    )

    class Meta:
        verbose_name = "Étudiant"
        verbose_name_plural = "Étudiants"
        ordering = ['numero']
        unique_together = ['pv', 'matricule']

    def __str__(self):
        return f"{self.numero} - {self.nom_prenom} ({self.matricule})"

    def get_decision_badge_class(self):
        """Retourne la classe CSS Bootstrap pour le badge de décision"""
        return {
            'V': 'bg-success',
            'NV': 'bg-danger',
            'VC': 'bg-warning text-dark',
        }.get(self.decision_generale, 'bg-secondary')

    def get_decision_icon(self):
        """Retourne l'icône Bootstrap Icons pour la décision"""
        return {
            'V': 'bi-check-circle-fill',
            'NV': 'bi-x-circle-fill',
            'VC': 'bi-exclamation-circle-fill',
        }.get(self.decision_generale, 'bi-question-circle-fill')

    def calculer_moyenne_generale(self):
        """
        Calcule la moyenne générale de l'étudiant à partir de ses notes ECUE.
        Moyenne pondérée par les crédits de chaque ECUE.
        Retourne None si aucune note n'est disponible.
        """
        from django.db.models import Sum, F
        from decimal import Decimal

        # Récupérer toutes les notes de l'étudiant avec moyenne non nulle
        notes_valides = self.notes.filter(moyenne__isnull=False).select_related('ecue')

        if not notes_valides.exists():
            return None

        # Calcul de la somme pondérée et des crédits totaux
        somme_ponderee = Decimal('0')
        credits_totaux = 0

        for note in notes_valides:
            if note.moyenne is not None and note.ecue.credits:
                somme_ponderee += note.moyenne * note.ecue.credits
                credits_totaux += note.ecue.credits

        if credits_totaux == 0:
            return None

        # Calculer la moyenne pondérée
        moyenne = somme_ponderee / credits_totaux

        # Arrondir à 2 décimales
        return round(moyenne, 2)

    def calculer_credits_acquis(self):
        """
        Calcule le nombre total de crédits acquis par l'étudiant.
        Somme les crédits des ECUE validées (moyenne >= 10 ou décision = 'V' ou 'VC').
        Retourne 0 si aucun crédit acquis.
        """
        from django.db.models import Sum, Q

        # Compter les crédits des notes validées
        credits_notes = self.notes.filter(
            Q(moyenne__gte=10) | Q(decision__in=['V', 'VC'])
        ).select_related('ecue').aggregate(
            total=Sum('ecue__credits')
        )['total'] or 0

        return credits_notes

    def determiner_decision(self):
        """
        Détermine la décision générale de l'étudiant selon les règles académiques:
        - V (Validé): moyenne >= 10 ET tous les crédits acquis
        - VC (Validé par Compensation): moyenne >= 10 ET crédits partiels
        - NV (Non Validé): moyenne < 10
        Retourne None si la moyenne n'a pas pu être calculée.
        """
        from decimal import Decimal

        # Calculer la moyenne si elle n'existe pas
        if self.moyenne_generale is None:
            moyenne = self.calculer_moyenne_generale()
        else:
            moyenne = self.moyenne_generale

        if moyenne is None:
            return None

        # Calculer les crédits totaux du semestre
        credits_totaux_semestre = self.pv.ues.aggregate(
            total=Sum('ecues__credits')
        )['total'] or 0

        # Calculer les crédits acquis
        credits = self.calculer_credits_acquis()

        # Appliquer les règles de décision
        if moyenne >= Decimal('10'):
            if credits >= credits_totaux_semestre:
                return 'V'  # Validé - tous les crédits acquis
            else:
                return 'VC'  # Validé par Compensation - crédits partiels
        else:
            return 'NV'  # Non Validé - moyenne insuffisante

    def mettre_a_jour_resultats(self):
        """
        Met à jour la moyenne générale, les crédits acquis et la décision de l'étudiant.
        Utilise les méthodes de calcul automatique.
        Sauvegarde l'instance après mise à jour.
        """
        from django.db.models import Sum

        self.moyenne_generale = self.calculer_moyenne_generale()
        self.credits_acquis = self.calculer_credits_acquis()
        self.decision_generale = self.determiner_decision()
        self.save()


class Note(models.Model):
    """
    Modèle représentant une note d'un étudiant dans une ECUE
    """
    DECISION_CHOICES = [
        ('V', 'Validé'),
        ('NV', 'Non Validé'),
        ('VC', 'Validé par Compensation'),
    ]

    etudiant = models.ForeignKey(Etudiant, on_delete=models.CASCADE, related_name='notes')
    ecue = models.ForeignKey(ECUE, on_delete=models.CASCADE, related_name='notes')
    cc = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        validators=[MinValueValidator(0), MaxValueValidator(20)],
        verbose_name="Contrôle Continu",
        null=True,
        blank=True
    )
    examen = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        validators=[MinValueValidator(0), MaxValueValidator(20)],
        verbose_name="Examen",
        null=True,
        blank=True
    )
    moyenne = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        validators=[MinValueValidator(0), MaxValueValidator(20)],
        verbose_name="Moyenne",
        null=True,
        blank=True
    )
    credit_attribue = models.IntegerField(
        verbose_name="Crédit attribué",
        null=True,
        blank=True
    )
    decision = models.CharField(
        max_length=10,
        choices=DECISION_CHOICES,
        verbose_name="Décision",
        null=True,
        blank=True
    )

    class Meta:
        verbose_name = "Note"
        verbose_name_plural = "Notes"
        ordering = ['ecue__ordre']
        unique_together = ['etudiant', 'ecue']

    def __str__(self):
        return f"{self.etudiant.nom_prenom} - {self.ecue.code}: {self.moyenne}/20"

    def get_decision_badge_class(self):
        """Retourne la classe CSS Bootstrap pour le badge de décision"""
        return {
            'V': 'bg-success',
            'NV': 'bg-danger',
            'VC': 'bg-warning text-dark',
        }.get(self.decision, 'bg-secondary')


class SyntheseUE(models.Model):
    """
    Modèle représentant la synthèse d'un étudiant pour une UE
    """
    DECISION_CHOICES = [
        ('V', 'Validé'),
        ('NV', 'Non Validé'),
        ('VC', 'Validé par Compensation'),
    ]

    etudiant = models.ForeignKey(Etudiant, on_delete=models.CASCADE, related_name='syntheses_ue')
    ue = models.ForeignKey(UE, on_delete=models.CASCADE, related_name='syntheses')
    moyenne_ue = models.DecimalField(
        max_digits=5,
        decimal_places=2,
        validators=[MinValueValidator(0), MaxValueValidator(20)],
        verbose_name="Moyenne UE",
        null=True,
        blank=True
    )
    credits_attribues = models.IntegerField(
        verbose_name="Crédits attribués",
        null=True,
        blank=True
    )
    decision = models.CharField(
        max_length=10,
        choices=DECISION_CHOICES,
        verbose_name="Décision",
        null=True,
        blank=True
    )

    class Meta:
        verbose_name = "Synthèse UE"
        verbose_name_plural = "Synthèses UE"
        ordering = ['ue__ordre']
        unique_together = ['etudiant', 'ue']

    def __str__(self):
        return f"{self.etudiant.nom_prenom} - {self.ue.code}: {self.moyenne_ue}/20"
    
    def get_decision_badge_class(self):
        """Retourne la classe CSS pour le badge de décision"""
        return {
            'V': 'bg-success',
            'NV': 'bg-danger',
            'VC': 'bg-warning text-dark',
        }.get(self.decision, 'bg-secondary')
