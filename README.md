# AccountTester

## Description
AccountTester est une application Windows Forms (C#) permettant de tester divers aspects des comptes utilisateurs sur un système. Elle permet d'effectuer des vérifications sur la connectivité Internet, les droits d'accès aux lecteurs réseaux, la présence et les permissions d'Office, ainsi que la disponibilité des imprimantes.

## Features
- **Test de connexion Internet** : Vérifie si l'ordinateur a accès à Internet en envoyant une requête à Google.
- **Test des droits sur les lecteurs réseau** : Tente de créer et de supprimer un fichier test sur chaque lecteur réseau pour vérifier les permissions d'écriture.
- **Détection de la version d'Office** : Recherche la présence d'Office sur le système via la base de registre.
- **Test des droits de lecture et écriture Office** : Crée, modifie et lit un document Word pour vérifier les permissions de l'utilisateur.
- **Liste des imprimantes installées** : Affiche toutes les imprimantes disponibles sur le système.

### Features in development
| Nom | Desc. |
|---|---|
| **Rapport détaillé des tests** | Ajout d'un export des résultats sous format CSV ou JSON. | 
| **Interface améliorée** | Amélioration de l'UI pour une meilleure lisibilité des résultats. |
| **Support multi-utilisateur** | Permet de tester plusieurs comptes en une seule session. |

## Prerequisites
Avant d'exécuter le projet, assurez-vous d'avoir les éléments suivants installés :

- Windows avec .NET Framework installé.
- Microsoft Office installé (pour les tests relatifs à Word).
- Droits d'accès suffisants pour tester les lecteurs réseaux et la base de registre. 

## Usage
1. Ouvrir l'application.
2. Cliquer sur le bouton **Start**.
3. Attendre la fin des tests.
4. Consulter les résultats dans la zone de logs.

## Contributing
Les contributions sont les bienvenues ! Pour contribuer à ce projet, veuillez suivre ces étapes :

1. Forkez le dépôt.
2. Créez une nouvelle branche pour votre fonctionnalité (`git checkout -b my-new-feature`).
3. Apportez vos modifications.
4. Commitez vos changements (`git commit -m 'Add my new feature'`).
5. Poussez votre branche (`git push origin my-new-feature`).
6. Ouvrez une Pull Request.

## Issues and Suggestions
Si vous rencontrez des problèmes ou avez des suggestions pour améliorer le projet, veuillez utiliser le [GitHub issue tracker](https://github.com/Miiraak/Account-Tester/issues).

## License
Ce projet n'est pas licencé. Voir le fichier [LICENSE](./LICENSE) pour plus de détails.

## Authors
- [**Miiraak**](https://github.com/miiraak) - *Lead Developer*

