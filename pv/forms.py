from django import forms
from .models import ProcesVerbal


class PVUploadForm(forms.ModelForm):
    """
    Formulaire pour l'upload d'un fichier PV Excel
    """

    class Meta:
        model = ProcesVerbal
        fields = ['fichier']
        widgets = {
            'fichier': forms.FileInput(attrs={
                'class': 'form-control',
                'accept': '.xlsx,.xls',
                'id': 'pv-file-input'
            })
        }

    def clean_fichier(self):
        fichier = self.cleaned_data.get('fichier')

        if fichier:
            # Vérifier l'extension
            if not fichier.name.endswith(('.xlsx', '.xls')):
                raise forms.ValidationError("Le fichier doit être au format Excel (.xlsx ou .xls)")

            # Vérifier la taille (max 10 MB)
            if fichier.size > 10 * 1024 * 1024:
                raise forms.ValidationError("Le fichier ne doit pas dépasser 10 MB")

        return fichier
