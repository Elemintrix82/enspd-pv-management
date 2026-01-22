from django.contrib import admin
from django.utils.html import format_html
from .models import ProcesVerbal, Etudiant, UE, ECUE, Note, SyntheseUE

# Import Export (optionnel)
try:
    from import_export.admin import ImportExportModelAdmin
    HAS_IMPORT_EXPORT = True
except ImportError:
    ImportExportModelAdmin = admin.ModelAdmin
    HAS_IMPORT_EXPORT = False


@admin.register(ProcesVerbal)
class ProcesVerbalAdmin(ImportExportModelAdmin if HAS_IMPORT_EXPORT else admin.ModelAdmin):
    list_display = ['id', 'filiere_display', 'niveau', 'semestre', 'annee_academique', 'formation_badge', 'date_import', 'stats_display']
    list_filter = ['filiere', 'niveau', 'semestre', 'formation', 'date_import']
    search_fields = ['filiere', 'annee_academique']
    readonly_fields = ['date_import', 'stats_detail']

    def filiere_display(self, obj):
        return format_html(
            '<strong style="color: #0066CC;">{}</strong>',
            obj.filiere
        )
    filiere_display.short_description = 'Filière'

    def formation_badge(self, obj):
        if obj.formation:
            color = '#217346' if 'ALT' in obj.formation.upper() else '#6c757d'
            return format_html(
                '<span style="background-color: {}; color: white; padding: 3px 8px; border-radius: 3px; font-size: 11px;">{}</span>',
                color, obj.formation
            )
        return '-'
    formation_badge.short_description = 'Formation'

    def stats_display(self, obj):
        return format_html(
            '<span style="color: green;">✓{}</span> / '
            '<span style="color: red;">✗{}</span> / '
            '<span style="color: orange;">~{}</span>',
            obj.nombre_valides, obj.nombre_non_valides, obj.nombre_valides_compensation
        )
    stats_display.short_description = 'Stats (V/NV/VC)'

    def stats_detail(self, obj):
        if obj.pk:
            return format_html(
                '<div style="padding: 15px; background: #f8f9fa; border-radius: 5px;">'
                '<h3 style="color: #0066CC;">Statistiques</h3>'
                '<p><strong>Total étudiants:</strong> {}</p>'
                '<p><strong>Validés:</strong> <span style="color: green;">{} ({}%)</span></p>'
                '<p><strong>Non Validés:</strong> <span style="color: red;">{}</span></p>'
                '<p><strong>Par Compensation:</strong> <span style="color: orange;">{}</span></p>'
                '<p><strong>Taux de réussite:</strong> <strong style="color: #0066CC; font-size: 18px;">{}%</strong></p>'
                '</div>',
                obj.nombre_etudiants,
                obj.nombre_valides,
                round(obj.nombre_valides / obj.nombre_etudiants * 100, 1) if obj.nombre_etudiants > 0 else 0,
                obj.nombre_non_valides,
                obj.nombre_valides_compensation,
                obj.taux_reussite
            )
        return "Sauvegardez d'abord"
    stats_detail.short_description = 'Détails Statistiques'


@admin.register(UE)
class UEAdmin(ImportExportModelAdmin if HAS_IMPORT_EXPORT else admin.ModelAdmin):
    list_display = ['code_display', 'intitule', 'pv', 'ordre', 'nb_ecues']
    list_filter = ['pv']
    search_fields = ['code', 'intitule']

    def code_display(self, obj):
        return format_html(
            '<code style="background: #e7f3ff; padding: 4px 8px; border-radius: 3px;">{}</code>',
            obj.code
        )
    code_display.short_description = 'Code'

    def nb_ecues(self, obj):
        count = obj.ecues.count()
        return format_html(
            '<span style="background: #217346; color: white; padding: 3px 8px; border-radius: 50%; font-size: 11px;">{}</span>',
            count
        )
    nb_ecues.short_description = 'ECUE'


@admin.register(ECUE)
class ECUEAdmin(ImportExportModelAdmin if HAS_IMPORT_EXPORT else admin.ModelAdmin):
    list_display = ['code_display', 'intitule_short', 'ue', 'ordre', 'credits_badge']
    list_filter = ['ue__pv', 'ue', 'credits']
    search_fields = ['code', 'intitule']

    def code_display(self, obj):
        return format_html(
            '<code style="background: #fff3cd; padding: 4px 8px; border-radius: 3px; font-weight: bold;">{}</code>',
            obj.code
        )
    code_display.short_description = 'Code'

    def intitule_short(self, obj):
        return obj.intitule[:50] + '...' if len(obj.intitule) > 50 else obj.intitule
    intitule_short.short_description = 'Intitulé'

    def credits_badge(self, obj):
        return format_html(
            '<span style="background: #0066CC; color: white; padding: 3px 10px; border-radius: 3px;">{}</span>',
            obj.credits
        )
    credits_badge.short_description = 'Crédits'


@admin.register(Etudiant)
class EtudiantAdmin(ImportExportModelAdmin if HAS_IMPORT_EXPORT else admin.ModelAdmin):
    list_display = ['numero', 'matricule_display', 'nom_prenom', 'moyenne_display', 'credits_acquis', 'decision_badge', 'pv']
    list_filter = ['decision_generale', 'pv']
    search_fields = ['matricule', 'nom_prenom']
    ordering = ['numero']

    def matricule_display(self, obj):
        return format_html(
            '<code style="background: #f8f9fa; padding: 4px 8px;">{}</code>',
            obj.matricule
        )
    matricule_display.short_description = 'Matricule'

    def moyenne_display(self, obj):
        color = 'green' if obj.moyenne_generale >= 10 else 'red'
        return format_html(
            '<strong style="color: {};">{}/20</strong>',
            color, obj.moyenne_generale
        )
    moyenne_display.short_description = 'Moyenne'

    def decision_badge(self, obj):
        colors = {
            'V': ('#28a745', 'white', '✓ Validé'),
            'NV': ('#dc3545', 'white', '✗ Non Validé'),
            'VC': ('#ffc107', 'black', '~ Par Compensation'),
        }
        bg, fg, text = colors.get(obj.decision_generale, ('#6c757d', 'white', obj.decision_generale))
        return format_html(
            '<span style="background: {}; color: {}; padding: 4px 12px; border-radius: 3px; font-weight: bold;">{}</span>',
            bg, fg, text
        )
    decision_badge.short_description = 'Décision'


@admin.register(Note)
class NoteAdmin(ImportExportModelAdmin if HAS_IMPORT_EXPORT else admin.ModelAdmin):
    list_display = ['etudiant_display', 'ecue_display', 'cc', 'examen', 'moyenne_display', 'credit_attribue', 'decision_badge']
    list_filter = ['decision', 'ecue']
    search_fields = ['etudiant__nom_prenom', 'etudiant__matricule', 'ecue__code']

    def etudiant_display(self, obj):
        return format_html(
            '<strong>{}</strong><br><small style="color: #6c757d;">{}</small>',
            obj.etudiant.nom_prenom[:30], obj.etudiant.matricule
        )
    etudiant_display.short_description = 'Étudiant'

    def ecue_display(self, obj):
        return format_html(
            '<code style="background: #e7f3ff; padding: 2px 6px;">{}</code>',
            obj.ecue.code
        )
    ecue_display.short_description = 'ECUE'

    def moyenne_display(self, obj):
        color = 'green' if obj.moyenne >= 10 else 'red'
        return format_html(
            '<strong style="color: {};">{}/20</strong>',
            color, obj.moyenne
        )
    moyenne_display.short_description = 'Moyenne'

    def decision_badge(self, obj):
        colors = {
            'V': '#28a745',
            'NV': '#dc3545',
            'VC': '#ffc107',
        }
        bg = colors.get(obj.decision, '#6c757d')
        return format_html(
            '<span style="background: {}; color: white; padding: 2px 8px; border-radius: 3px; font-size: 11px;">{}</span>',
            bg, obj.decision
        )
    decision_badge.short_description = 'Décision'


@admin.register(SyntheseUE)
class SyntheseUEAdmin(ImportExportModelAdmin if HAS_IMPORT_EXPORT else admin.ModelAdmin):
    list_display = ['etudiant_display', 'ue_display', 'moyenne_display', 'credits_attribues', 'decision_badge']
    list_filter = ['decision', 'ue']
    search_fields = ['etudiant__nom_prenom', 'etudiant__matricule', 'ue__code']

    def etudiant_display(self, obj):
        return format_html(
            '<strong>{}</strong>',
            obj.etudiant.nom_prenom[:30]
        )
    etudiant_display.short_description = 'Étudiant'

    def ue_display(self, obj):
        return format_html(
            '<code style="background: #e7f3ff; padding: 3px 8px;">{}</code>',
            obj.ue.code
        )
    ue_display.short_description = 'UE'

    def moyenne_display(self, obj):
        color = 'green' if obj.moyenne_ue >= 10 else 'red'
        return format_html(
            '<strong style="color: {};">{}/20</strong>',
            color, obj.moyenne_ue
        )
    moyenne_display.short_description = 'Moyenne UE'

    def decision_badge(self, obj):
        colors = {
            'V': '#28a745',
            'NV': '#dc3545',
            'VC': '#ffc107',
        }
        bg = colors.get(obj.decision, '#6c757d')
        return format_html(
            '<span style="background: {}; color: white; padding: 3px 10px; border-radius: 3px;">{}</span>',
            bg, obj.decision
        )
    decision_badge.short_description = 'Décision'
