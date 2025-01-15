#!/usr/bin/env python3
"""\
Objet       lire des maquettes depuis un fichier excel ou csv ou texte et le convertir en arbres de maquettes Pégase

Entrée      fichiers sources contenant la définition d'une maquette ou entrée standard
Sortie      représentation JSON des maquettes trouvées dans les fichiers lus

Usage       maquettes-xl2json.py [-n code,code,...] [-b] [-d] [fichier_excel[:i:j:k:...]] [fichier_excel[:i:j:k...]] ...
  fichier   le nom du ou des fichiers à traiter, avec éventuellement l'index du ou des onglets ciblés
            (si non précisé, la cible est le premier onglet)

  -a        affiche une aide de type 'usage' consistant en les présentes lignes
  -n        liste des codes à renvoyer au format JSON, séparés par une virgule - si non présent, renvoie toutes les racines trouvées
  -b        renvoie les maquettes encodées en base64
  -l        si présente, la présence des libellés n'est plus obligatoire (un libellé type sera généré automatiquement)
  -d        affiche des messages d'info pour suivre le déroulé de l'execution de la commande
  -g        si présente, la construction de maquette échoue dès lors qu'un groupement est spécifié sans plage de choix
  -h        pour les cas désespérés : transmet les indices de colonnes (séparés par virgule) où se trouvent les données
            type d'objet, nature, code, libellé, libellé long, ects, code du noeud parent, plage min de groupement, plage max

Auteur
Alfredo Pereira - 09/24
alfredo.pereira@inalco.fr
"""

usage="""
Usage       {} [-n code,code,...] [-b] [-d] [fichier_excel[:i:j:k:...]] [fichier_excel[:i:j:k...]] ...
  fichier   le nom du ou des fichiers à traiter, avec éventuellement l'index du ou des onglets ciblés
            (si non précisé, la cible est le premier onglet)

  -a        affiche une aide de type 'usage' consistant en les présentes lignes
  -n        liste des codes à renvoyer au format JSON, séparés par une virgule - si non présent, renvoie toutes les racines trouvées
  -b        renvoie les maquettes encodées en base64
  -l        si présente, la présence des libellés n'est plus obligatoire (un libellé type sera généré automatiquement)
  -d        affiche des messages d'info pour suivre le déroulé de l'execution de la commande
  -g        si présente, la construction de maquette échoue dès lors qu'un groupement est spécifié sans plage de choix
  -h        pour les cas désespérés : transmet les indices de colonnes (séparés par virgule) où se trouvent les données
            type d'objet, nature, code, libellé, libellé long, ects, code du noeud parent, plage min de groupement, plage max
"""

import sys
import json
import uuid
import zlib
import base64
import getopt
import fileinput
from pathlib import Path
from python_calamine import CalamineWorkbook

######################
# Variables globales #
######################

#
# Valeurs par défaut des paramètres de la commande
#
b64 = False             # correspond à l'option -b
msgs = False            # correspond à l'option -d
noeuds_demandes = []    # correspond à l'option -n
headers_param = dict()  # correspond à l'option -h
verif_choix_groupements = False        # option -g

#
# Mapping des titres de colonnes dans un excel/csv avec les attributs d'un objet de la classe 'NoeudMaquette'
#
donnees_csv = {
    'Type objet': 'type_noeud',
    'Code objet': 'code',
    'Libellé': 'libelle',
    'Libellé long': 'libelle_long',
    'Nature objet': 'nature',
    'ECTS objet': 'ects',
    'Plage MIN': 'plage_min',
    'Plage MAX': 'plage_max',
    'Code parent': 'code_parent',
    'Obligatoire': 'obligatoire_parent',
    'Mutualisé': 'est_mutualise',
    'Distanciel': 'est_distanciel',
    'Stage': 'est_stage',
    'Capacité accueil': 'capacite_accueil',
    'Structure principale': 'structure_principale',
    'ID objet': 'id_noeud',
    'Type formation': 'type_formation',
    'Syllabus - objectifs': 'syll_objectifs',
    'Syllabus - description': 'syll_description',
    'Syllabus - ouverture mobilité entrante': 'syll_ouverture_mobilite_entrante',
    'Syllabus - langue enseignement': 'syll_langue_enseignement',
    'Syllabus - prérequis': 'syll_prerequis_pedagogiques',
    'Syllabus - bibliographie': 'syll_bibliographie',
    'Syllabus - contacts': 'syll_contacts',
    'Syllabus - autres informations': 'syll_autres_infos',
    'Syllabus - modalités enseignement': 'syll_modalites_enseignement',
    'Syllabus - volume horaire': 'syll_volume_horaire',
    'Syllabus - coefficient': 'syll_coefficient',
    'Syllabus - modalités évaluation': 'syll_modalites_eval',
    'SISE - type diplôme': 'sise_type_diplome',
    'SISE - code diplôme': 'sise_code_diplome',
    'SISE - niveau diplôme SISE': 'sise_niveau_diplome_sise',
    'SISE - parcours-type': 'sise_parcours_type',
    'SISE - domaine formation': 'sise_domaine_formation',
    'SISE - mention': 'sise_mention',
    'SISE - champ formation': 'sise_champ_formation',
    'SISE - niveau diplôme': 'sise_niveau_diplome',
    'SISE - déclinaison': 'sise_declinaison',
    'Aglae - habilité bourses': 'aglae_habilite_bourses',
    'Aglae - niveau': 'aglae_niveau',
    'Formation porteuse': 'formation_porteuse',
    'Structures porteuses': 'structures_porteuses'
}

#
# Liste des noms de colonnes excel/csv/texte obligatoires
#
donnees_csv_obligatoires = ['Type objet', 'Code objet', 'Libellé']

#
# Valeurs par défaut d'une ligne de données lue dans un fichier ou sur l'entrée standard
#
noeud_defaults = {
    'type_noeud': None,
    'code': None,
    'libelle': None,
    'libelle_long': None,
    'nature': None,
    'ects': None,
    'plage_min': None,
    'plage_max': None,
    'code_parent': None,
    'obligatoire_parent': True,
    'est_mutualise': False,
    'est_distanciel': False,
    'est_stage': False,
    'capacite_accueil': None,
    'structure_principale': None,
    'id_noeud': None,
    'type_formation': '0',
    'syll_objectifs': None,
    'syll_description': None,
    'syll_ouverture_mobilite_entrante': None,
    'syll_langue_enseignement': None,
    'syll_prerequis_pedagogiques': None,
    'syll_bibliographie': None,
    'syll_contacts': None,
    'syll_autres_infos': None,
    'syll_modalites_enseignement': None,
    'syll_volume_horaire': None,
    'syll_coefficient': None,
    'syll_modalites_eval': None,
    'sise_type_diplome': None,
    'sise_code_diplome': None,
    'sise_niveau_diplome_sise': None,
    'sise_parcours_type': None,
    'sise_domaine_formation': None,
    'sise_mention': None,
    'sise_champ_formation': None,
    'sise_niveau_diplome': None,
    'sise_declinaison': None,
    'aglae_habilite_bourses': False,
    'aglae_niveau': None,
    'formation_porteuse': None,
    'structures_porteuses': None
}

#
# Mapping de qq valeurs booléennes qu'on peut trouver dans un excel/csv
#
bool_equiv = {
    'oui'   : True,
    'non'   : False,
    'false' : False,
    'true'  : True,
    'null'  : None
}


def process_line(ligne, headers_courants):
    """Traiter une ligne de fichier spécifiant les données d'un noeud de maquette en tant que liste"""

    #
    # Utiliser en priorité les headers fournis dans la commande, si présents
    #
    if headers_param:
        headers_courants = headers_param
    else:
        #
        # Tester si la ligne courante est une ligne de headers - critère : la ligne contient les libellés des données obligatoires
        #
        if [True for d in donnees_csv_obligatoires if d in ligne] == [True] * len(donnees_csv_obligatoires):
            if msgs: print('Détection d\'une ligne de header', file=sys.stderr)

            headers_courants.clear()

            #
            # Construction de l'index des données se trouvant dans le fichier source
            #
            for i, x in enumerate(ligne):
                if donnees_csv.get(x): headers_courants[donnees_csv[x]] = i

            if msgs: print('Colonnes détectées :', headers_courants)
            return

    if not headers_courants:
        if msgs: print(ligne, ': ligne ignorée car aucun header n\'a été défini dans cette passe', file=sys.stderr)
        return

    #
    # Chargement des valeurs par défaut d'un objet NoeudMaquette
    #
    valeurs_noeud = dict(noeud_defaults)

    #
    # Mise à jour de la variable valeurs_noeud avec, outre les valeurs par défaut, les valeurs trouvées dans la ligne de données courante
    #
    for h in headers_courants:
        try:
            if ligne[headers_courants[h]] != '': valeurs_noeud[h] = str(ligne[headers_courants[h]])
        except:
            # if msgs: print('Pas de donnée', h, 'trouvée dans', ligne, file=sys.stderr)
            pass

        try:
            valeurs_noeud[h] = bool_equiv[valeurs_noeud[h].lower()]
        except:
            pass

    #
    # Cette portion de code (contrôle de cohérence) serait mieux située dans l'initialisation d'un objet NoeudMaquette --> Plus tard
    #
    type_noeud = valeurs_noeud.get('type_noeud')

    if type_noeud:
        type_noeud = type_noeud.upper()
    else:
        if msgs: print(ligne, ': ligne ignorée car pas de type d\'objet de formation indiqué', file=sys.stderr)
        return

    #
    # Création, en fonction du type d'objet de formation indiqué, d'une instance de la classe correcte
    #
    try:
        if type_noeud == 'FORMATION':
            noeud = NoeudFormation(valeurs_noeud)
        elif type_noeud == 'GROUPEMENT':
            noeud = NoeudGroupement(valeurs_noeud)
        else:
            noeud = NoeudObjetFormation(valeurs_noeud)

    except ValueError as erreur:
        if msgs: print(erreur, file=sys.stderr)
        if 'plages de choix incomplètes' in str(erreur): sys.exit(1)



class NoeudMaquette:
    #
    # Longueurs maximales de champs critiques
    #
    lg_max_code, lg_max_libelle, lg_max_libelle_long = 30, 50, 150

    #
    # Dictionnaire des noeuds créés jusqu'ici, indexés par leur code
    #
    noeuds = dict()


    def __init__(self, val):
        #
        # Vérification de la présence des données obligatoires pour créer un noeud
        #
        for d in donnees_csv_obligatoires:
            if not val.get(donnees_csv[d]): raise ValueError('Donnée obligatoire manquante : ' + d)

        #
        # Ajustement des paramètres passés dans la variable 'val'
        #
        if not val['code'] or len(val['code']) > self.lg_max_code:
            # Erreur sur le code, on ne peut (et ne doit) rien faire
            raise ValueError('Erreur sur le code ' + str(val['code']))

        #
        # Les codes en majuscules
        #
        val['code'] = val['code'].upper()
        if val['code_parent']: val['code_parent'] = val['code_parent'].upper()

        #
        # Convertir en nombre à virgule les ects et en entiers les plages des groupements
        #
        try:
            val['ects'] = float(val['ects'])
        except:
            val['ects'] = None

        try:
            val['plage_max'] = int(val['plage_max'])
        except:
            val['plage_max'] = None

        try:
            val['plage_min'] = int(val['plage_min'])
        except:
            val['plage_min'] = None

        #
        # Libellé automatique si l'option -l est activée
        #
        if not val['libelle']:
            val['libelle'] = 'Objet de type ' + val['type_noeud'] + ' et de code ' + val['code']

        #
        # Dupliquer le libellé si le libellé long est absent des données
        #
        if not val['libelle_long']:
            val['libelle_long'] = val['libelle']

        #
        # Contrôle des longueurs de libellés
        #
        if len(val['libelle']) > self.lg_max_libelle:
            # Tronquer à la longueur max si le libellé est trop long
            val['libelle'] = val['libelle'][:self.lg_max_libelle]

            if msgs: print(val['code'], ': libellé trop long, tronqué à', self.lg_max_libelle, file=sys.stderr)

        if len(val['libelle_long']) > self.lg_max_libelle_long:
            # Tronquer à la longueur max si le libellé long dépasse la longueur autorisée
            val['libelle_long'] = val['libelle_long'][:self.lg_max_libelle_long]

            if msgs: print(val['code'], ': libellé long trop long, tronqué à', self.lg_max_libelle_long, file=sys.stderr)

        #
        # Assignation d'un uuid aléatoire si aucun uuid n'est fourni en donnée
        #
        if not val['id_noeud']:
            val['id_noeud'] = str(uuid.uuid4()).lower()



        if msgs: print('Traitement du noeud', val['code'], file=sys.stderr)

        #
        # Le code fourni en donnée n'est pas encore apparu --> Création d'un nouveau noeud
        #
        if not val['code'] in NoeudMaquette.noeuds:
            #
            # Création des membres communs de la classe NoeudMaquette
            #
            self.type_noeud = val['type_noeud']
            self.id = val['id_noeud']
            self.code = val['code']
            self.mutualise = val['est_mutualise']
            self.type = None

            #
            # Les contextes : fonctionnalité à ajouter dans une future version
            #
            self.contextes = []

            self.descripteursObjetMaquette = {
                'libelle': val['libelle'],
                'libelleLong': val['libelle_long']
            }

            #
            # Création éventuelle de la liste des structures porteuses
            #
            val['structures_porteuses'] = val['structures_porteuses'].split(',') if val['structures_porteuses'] else []

            self.formatsEnseignement = {
                'structuresPorteuse': val['structures_porteuses'],
                # 'formationPorteuse': val['formation_porteuse'],
                'formatsEnseignement': []
            }

            #
            # Ensemble des noeuds enfants, pour l'instant vide puisque création de nouveau noeud
            #
            self.enfants = set()

            #
            # Ensemble des ascendants, nécessaire pour éviter les références circulaires
            #
            self.ascendants = set()


        #
        # Exclusions de cas où la donnée fournie n'est pas cohérente
        #

        # Le code parent indiqué n'existe pas
        if val['code_parent'] and val['code_parent'] not in NoeudMaquette.noeuds:
            raise ValueError('Problème avec ' + str(val['code']) + ', le code parent indiqué n\'existe pas : ' + str(val['code_parent']))


        # Le code existe déjà mais pas de code parent fourni --> rien à faire (pas d'update de noeuds)
        if val['code'] in NoeudMaquette.noeuds and not val['code_parent']:
            raise ValueError('Noeud déjà traité, sans indication de nouveau parent : ' + str(val['code']))


        # Le code indiqué est déjà enfant du code parent fourni --> rien à faire (pas d'update de noeud)
        if val['code_parent'] in NoeudMaquette.noeuds and val['code'] in NoeudMaquette.noeuds[val['code_parent']].enfants:
            raise ValueError(str(val['code']) + ' est déjà enfant de ' + str(val['code_parent']))



        #
        # Pas d'anomalie de valeurs détectée, on peut créer un lien de parenté entre les 2 noeuds fournis (code et code_parent)
        #

        if val['code_parent'] in NoeudMaquette.noeuds:

            # Le noeud créé a un code déjà rencontré
            if val['code'] in NoeudMaquette.noeuds:
                try:
                    NoeudMaquette.creer_enfant(NoeudMaquette.noeuds[val['code_parent']], NoeudMaquette.noeuds[val['code']])
                except ValueError as erreur:
                    raise ValueError(erreur)

                # Pas vraiment une "erreur", mais l'initialisation doit cesser ici et être remontée en tant qu'exception
                raise ValueError('Le noeud ' + str(val['code']) + ' existe déjà, ajout en tant qu\'enfant de ' + str(val['code_parent']))

            # Le noeud créé est un nouveau noeud
            else:
                try:
                    NoeudMaquette.creer_enfant(NoeudMaquette.noeuds[val['code_parent']], self)
                except ValueError as erreur:
                    raise ValueError(erreur)



    def __str__(self):
        #
        # Sérialisation avec json.dumps du dictionnaire représentant une instance d'objet NoeudMaquette
        #
        return  json.dumps(self, cls=NoeudMaquetteEncoder, separators=(',', ':'))


    def creer_enfant(parent, enfant):
        #
        # Création d'un lien parent-enfant entre deux noeuds
        #

        # Vérifier si pas de référence circulaire
        if enfant in NoeudMaquette.noeuds[parent.code].ascendants:
            raise ValueError('Le noeud ' + str(enfant.code) + ' ne peut devenir enfant de l\'un de ses descendants')

        parent.enfants.add(enfant)
        enfant.ascendants.add(parent)
        enfant.ascendants |= parent.ascendants



class NoeudMaquetteEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, NoeudObjetFormation) or isinstance(o, NoeudFormation):
            return {
                'id':           o.id,
                'code':         o.code,
                'mutualise':    o.mutualise,
                'type':         o.type,
                'contextes':    o.contextes,
                'descripteursObjetMaquette':    o.descripteursObjetMaquette,
                'descripteursSyllabus':         o.descripteursSyllabus,
                'descripteursEnquete':          o.descripteursEnquete,
                'formatsEnseignement':          o.formatsEnseignement,
                'enfants':                      [{'obligatoire':True,'objetMaquette':e} for e in o.enfants]
            }
        elif isinstance(o, NoeudMaquette):
            return {
                'id':           o.id,
                'code':         o.code,
                'mutualise':    o.mutualise,
                'type':         o.type,
                'contextes':    o.contextes,
                'descripteursObjetMaquette':    o.descripteursObjetMaquette,
                'descripteursEnquete':          o.descripteursEnquete,
                'formatsEnseignement':          o.formatsEnseignement,
                'enfants':                      [{'obligatoire':True,'objetMaquette':e} for e in o.enfants]
            }

        return super().default(o)


#
# Classe non utilisée pour le moment, elle servira à définir les contextes de noeuds
#
class ContexteNoeud:
    def __init__(
        self,
        val
    ):
        self.id = str(uuid.uuid4()).lower()
        self.chemin = [] # fabriquer le chemin en suite de codes
        self.valide = False

        if val['type_noeud'] == 'FORMATION':
            self.type = 'FormationContexteEntity'
        elif val['type_noeud'] == 'GROUPEMENT':
            self.type = 'GroupementContexteEntity'
        else:
            self.type = 'ObjetFormationContexteEntity'

        self.pointInscriptionAdministrative = {
            'inscriptionAdministrative': False,
            'actif': False
        }


class NoeudFormation(NoeudMaquette):
    def __init__(self, val):
        try:
            super().__init__(val)
        except ValueError as erreur:
            raise ValueError(erreur)

        self.type = 'FormationEntity'

        self.descripteursObjetMaquette.update({
            'ects': val['ects'],
            'structurePrincipale': val['structure_principale'],
            'teleEnseignement': val['est_distanciel'],
            'typeFormation': val['type_formation']
        })

        self.descripteursSyllabus = {
            'objectif': val['syll_objectifs'],
            'description':  val['syll_description'],
            'ouvertureALaMobiliteEntrante': val['syll_ouverture_mobilite_entrante'],
            'langueEnseignement':   val['syll_langue_enseignement'],
            'prerequisPedagogique': val['syll_prerequis_pedagogiques'],
            'bibliographie':    val['syll_bibliographie'],
            'contacts': val['syll_contacts'],
            'autresInformations':   val['syll_autres_infos'],
            'modalitesEnseignements':   val['syll_modalites_enseignement'],
            'volumeHoraireParTypeDeCours':  val['syll_volume_horaire'],
            'coefficient':  val['syll_coefficient'],
            'modalitesEvaluation':  val['syll_modalites_eval']
        }

        self.descripteursEnquete = {
            'enqueteAglae': {
                'habilitePourBoursesAglae': val['aglae_habilite_bourses'],
                'niveauAglae': val['aglae_niveau']
            },
            'enqueteSise': {
                'typeDiplome': val['sise_type_diplome'],
                'codeDiplomeSise': val['sise_code_diplome'],
                'niveauDiplomeSise': val['sise_niveau_diplome_sise'],
                'parcoursTypeSise': val['sise_parcours_type'],
                'domaineFormation': val['sise_domaine_formation'],
                'mention': val['sise_mention'],
                'champFormation': val['sise_champ_formation'],
                'niveauDiplome': val['sise_niveau_diplome'],
                'declinaisonDiplome': val['sise_declinaison']
            }
        }

        # Ajout du noeud nouvellement créé à l'ensemble des noeuds
        NoeudMaquette.noeuds[self.code] = self


class NoeudObjetFormation(NoeudMaquette):
    def __init__(self, val):
        try:
            super().__init__(val)
        except ValueError as erreur:
            raise ValueError(erreur)

        self.type = 'ObjetFormationEntity'

        self.descripteursObjetMaquette.update({
            'ects': val['ects'],
            'structurePrincipale': val['structure_principale'],
            'typeObjetFormation': val['type_noeud'],
            'nature': val['nature'],
            'stage': val['est_stage'],
            'teleEnseignement': val['est_distanciel'],
            'capaciteAccueil': val['capacite_accueil']
        })

        self.descripteursSyllabus = {
            'objectif': val['syll_objectifs'],
            'description':  val['syll_description'],
            'ouvertureALaMobiliteEntrante': val['syll_ouverture_mobilite_entrante'],
            'langueEnseignement':   val['syll_langue_enseignement'],
            'prerequisPedagogique': val['syll_prerequis_pedagogiques'],
            'bibliographie':    val['syll_bibliographie'],
            'contacts': val['syll_contacts'],
            'autresInformations':   val['syll_autres_infos'],
            'modalitesEnseignements':   val['syll_modalites_enseignement'],
            'volumeHoraireParTypeDeCours':  val['syll_volume_horaire'],
            'coefficient':  val['syll_coefficient'],
            'modalitesEvaluation':  val['syll_modalites_eval']
        }

        self.descripteursEnquete = {
            'enqueteAglae': {
                'habilitePourBoursesAglae': val['aglae_habilite_bourses'],
                'niveauAglae': val['aglae_niveau']
            },
            'enqueteSise': {
                'typeDiplome': val['sise_type_diplome'],
                'codeDiplomeSise': val['sise_code_diplome'],
                'niveauDiplomeSise': val['sise_niveau_diplome_sise'],
                'parcoursTypeSise': val['sise_parcours_type'],
                'domaineFormation': val['sise_domaine_formation'],
                'mention': val['sise_mention'],
                'champFormation': val['sise_champ_formation'],
                'niveauDiplome': val['sise_niveau_diplome'],
                'declinaisonDiplome': val['sise_declinaison']
            }
        }

        # Ajout du noeud nouvellement créé à l'ensemble des noeuds
        NoeudMaquette.noeuds[self.code] = self


class NoeudGroupement(NoeudMaquette):
    def __init__(self, val):
        try:
            super().__init__(val)
        except ValueError as erreur:
            raise ValueError(erreur)

        self.type = 'GroupementEntity'

        if val['plage_min'] and val['plage_max']:
            self.descripteursObjetMaquette.update({
                'nature': val['nature'],
                'plageDeChoix': {
                    'min': val['plage_min'],
                    'max': val['plage_max']
                }
            })
        else:
            if verif_choix_groupements:
                raise ValueError(val['code'] + ' : plages de choix incomplètes dans le groupement')

        self.descripteursEnquete = {
            'enqueteAglae': {
                'habilitePourBoursesAglae': val['aglae_habilite_bourses'],
                'niveauAglae': val['aglae_niveau']
            }
        }

        # Ajout du noeud nouvellement créé à l'ensemble des noeuds
        NoeudMaquette.noeuds[self.code] = self


def main():
    ############################################################
    # Traitement de la commande et de ses paramètres éventuels #
    ############################################################

    argv = sys.argv
    commande = argv[0]

    #
    # Parser les arguments de la commande avec le module getopt
    #
    try:
        opts, args = getopt.gnu_getopt(argv[1:], "an:bdh:lg")
    except:
        print(usage.format(commande).strip(), file=sys.stderr)
        sys.exit(1)


    #
    # Paramètres généraux de la commande
    #
    global b64, msgs, noeuds_demandes, headers_param, verif_choix_groupements, donnees_csv_obligatoires


    #
    # Analyse des options fournies en commande
    #
    for opt, arg in opts:
        if   opt == '-b':
            b64 = True

        elif opt == '-n':
            noeuds_demandes = [x.upper() for x in arg.split(',')]

        elif opt == '-d':
            msgs = True

        elif opt == '-l':
            donnees_csv_obligatoires = ['Type objet', 'Code objet']

        elif opt == '-g':
            verif_choix_groupements = True

        elif opt == '-h':
            param_h = []

            for i in arg.split(','):
                if i.isdigit():
                    param_h.append(int(i)-1)
                else:
                    param_h.append('')

            headers_param = dict(zip(['type_noeud', 'nature', 'code', 'libelle', 'libelle_long', 'ects', 'code_parent', 'plage_min', 'plage_max'], param_h))
            for delh in [h for h in headers_param if not isinstance(headers_param[h], int)]: del headers_param[delh]

        elif opt == '-a':
            print(usage.format(commande).strip())
            sys.exit(0)

        argv.remove(opt)
        if arg: argv.remove(arg)


    ###############################
    # Traitement des données lues #
    ###############################

    headers_courants = dict()

    if msgs:
        if headers_param:
            print('Le traitement du fichier se fera suivant', headers_param, file=sys.stderr)
        else:
            print('Le traitement du fichier se fera suivant les headers rencontrés dans les fichiers', file=sys.stderr)


    #
    # Si pas de fichier spécifié en commande, on se branche sur l'entrée standard
    #
    if not argv[1:]:
        if msgs: print('Lecture des données sur l\'entrée standard', file=sys.stderr)

        for ligne in sys.stdin:
            ligne = [l.strip() for l in ligne.split('\t')]
            process_line(ligne, headers_courants)

    else:
        #
        # Traitement des noms de fichiers spécifiés en argument de commande
        #
        fichiers = argv[1:]

        for ind, arg in enumerate(fichiers):
            #
            # Recherche d'éventuelles indications d'onglets
            #
            arg = arg.split(':')
            nom_fichier = arg.pop(0)

            #
            # A-t-on un répertoire en paramètre ? Si oui, on ajoute à la liste des fichiers à traiter les fichiers (ou répertoires) présents dans le répertoire en question
            #
            if Path(nom_fichier).is_dir():
                if msgs: print('Parcours du répertoire', nom_fichier, file=sys.stderr)

                if arg:
                    fichiers[ind+1:ind+1] = [str(f) + ':' + ':'.join(arg) for f in Path(nom_fichier).iterdir()]
                else:
                    fichiers[ind+1:ind+1] = [str(f) for f in Path(nom_fichier).iterdir()]

                continue


            if msgs: print('Lecture du fichier', nom_fichier, file=sys.stderr)

            #
            # Chercher l'extension du fichier pour déterminer son format --> texte, csv, excel
            #
            extension = Path(nom_fichier).suffix

            if extension in ['.txt', '.csv']:
                try:
                    fichier = open(nom_fichier, encoding='utf-8', errors="replace")
                except OSError:
                    print('Impossible d\'ouvrir le fichier', nom_fichier, file=sys.stderr)
                else:
                    #
                    # Lecture ligne à ligne d'un fichier texte ou csv
                    #
                    for ligne in fichier:
                        ligne = [l.strip() for l in ligne.split('\t')]                        
                        process_line(ligne, headers_courants)

                    fichier.close()

            #
            # Si l'extension commence par 'xl', supposons que c'est bien un fichier excel
            #
            elif extension[1:3].lower() == 'xl':
                try:
                    workbook = CalamineWorkbook.from_path(nom_fichier)
                except:
                    print('Impossible d\'ouvrir le fichier', nom_fichier, file=sys.stderr)
                else:
                    #
                    # Déterminer les onglets à traiter - soit ils sont indiqués par numéro d'index soit on spécifie le début de leur nom
                    # Si pas d'indication, traitement du premier onglet trouvé dans le document
                    #

                    onglets=workbook.sheet_names
                    onglets_cibles=[]

                    if not arg:
                        #
                        # Aucun séparateur n'a été indiqué après le nom de fichier --> on traite le 1er onglet trouvé
                        #
                        onglets_cibles = [onglets[0]]

                    else:
                        if not arg[0]:
                            #
                            # Un séparateur a été spécifié mais sans valeur --> on traite tous les onglets du document
                            #
                            onglets_cibles = onglets

                        else:
                            if arg[0].isdigit():
                                #
                                # Des numéros d'onglets ont été fournis --> construction de la liste des onglets à traiter
                                #
                                onglets_cibles = [onglets[int(i)-1] for i in arg if i.isdigit() and int(i)-1 in range(len(onglets))]

                            else:
                                #
                                # Une chaîne non numérique a été fournie --> traitement de tous les onglets commençant par ladite chaîne
                                #
                                onglets_cibles = [onglet for onglet in onglets if onglet.startswith(arg[0])]

                    #
                    # Traitement des onglets du fichier courant
                    #
                    if onglets_cibles:
                        if msgs: print(nom_fichier, ': les onglets traités sont', onglets_cibles, file=sys.stderr)

                        for onglet in onglets_cibles:
                            lignes = iter(workbook.get_sheet_by_name(onglet).to_python())

                            for ligne in lignes:
                                # Stripper les chaînes de caractères
                                ligne = list(map(lambda l: l.strip() if isinstance(l, str) else l, ligne))

                                # Convertir en chaîne de caractères les nombres (important si la ligne a été produite par calamine_python)
                                ligne = list(map(lambda x: str(int(x)) if isinstance(x, float) and x.is_integer() else str(x), ligne))

                                process_line(ligne, headers_courants)

                            #
                            # Remise à zéro des headers lorsque l'on change d'onglet
                            #
                            headers_courants.clear()

                    else:
                        if msgs: print(nom_fichier, ': aucun onglet à traiter', file=sys.stderr)


            else:
                print(nom_fichier, ': fichier non traité car ayant une extension non reconnue', file=sys.stderr)


    ################################
    # Fin de traitement, affichage #
    ################################

    #
    # Si pas d'option -n, affichage de tous les noeuds racines rencontrés dans les fichiers ou sur l'entrée standard
    #
    if noeuds_demandes:
        if len(noeuds_demandes)>1 or ':' not in noeuds_demandes[0]:
            noeuds_demandes = [n for n in noeuds_demandes if n in NoeudMaquette.noeuds]
        else:
            noeuds_demandes = noeuds_demandes[0].split(':')
            type_fonction = noeuds_demandes[0]
            fonc_demandes = noeuds_demandes[1:]

            if type_fonction == 'F': # F comme filtre
                if fonc_demandes[0] != '':
                    noeuds_demandes = [NoeudMaquette.noeuds[n].code for n in NoeudMaquette.noeuds if NoeudMaquette.noeuds[n].type_noeud in fonc_demandes]
            elif type_fonction == 'B': # B comme branche
                noeuds_demandes = []
                for branche in fonc_demandes:
                    if NoeudMaquette.noeuds.get(branche):
                        noeuds_demandes += branche
                        noeuds_demandes += [NoeudMaquette.noeuds[n].code for n in NoeudMaquette.noeuds if NoeudMaquette.noeuds[branche] in NoeudMaquette.noeuds[n].ascendants]
            else:
                noeuds_demandes = []
    else:
        #
        # Les racines sont des noeuds avec un ensemble d'ascendants vide, ie de cardinal zéro
        #
        noeuds_demandes = [NoeudMaquette.noeuds[n].code for n in NoeudMaquette.noeuds if len(NoeudMaquette.noeuds[n].ascendants) == 0]

    for n in noeuds_demandes:
        if b64:
            #
            # Compression de la donnée chargée (gzip) puis encodage en base 64 du résultat
            #
            compressor = zlib.compressobj(wbits=25)
            data = str(NoeudMaquette.noeuds[n]).encode()
            dataz = compressor.compress(data)
            dataz += compressor.flush()
            dataz = base64.b64encode(dataz).decode()
            print(dataz)

        else:
            print(NoeudMaquette.noeuds[n])

    #
    # Fin de main()
    #

if __name__ == '__main__':
    main()
