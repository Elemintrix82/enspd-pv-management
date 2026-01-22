from django import template
from django.http import QueryDict

register = template.Library()


@register.simple_tag(takes_context=True)
def url_replace(context, **kwargs):
    """
    Template tag pour remplacer/ajouter des paramètres GET
    Usage: {% url_replace page=2 %}
    """
    query = context['request'].GET.copy()
    for key, value in kwargs.items():
        query[key] = value
    return query.urlencode()


@register.simple_tag(takes_context=True)
def url_remove(context, *args):
    """
    Template tag pour supprimer des paramètres GET
    Usage: {% url_remove 'page' 'per_page' %}
    """
    query = context['request'].GET.copy()
    for key in args:
        if key in query:
            del query[key]
    return query.urlencode()


@register.simple_tag(takes_context=True)
def get_params_except(context, *exclude):
    """
    Retourne tous les paramètres GET sauf ceux spécifiés
    Usage: {% get_params_except 'page' %}
    """
    query = context['request'].GET.copy()
    for key in exclude:
        if key in query:
            del query[key]
    return query.urlencode()


@register.filter
def get_item(dictionary, key):
    """Récupère une valeur d'un dictionnaire par clé"""
    return dictionary.get(key)


# NOUVEAU : Filtre multiply
@register.filter
def multiply(value, arg):
    """Multiplie la valeur par l'argument"""
    try:
        return int(value) * int(arg)
    except (ValueError, TypeError):
        try:
            return value * arg
        except Exception:
            return ''