# automatisation-de-facturation-et-simulation-de-cr-dit-VBA-excel![facturation](https://github.com/HAMZA25032005/Automatisation-de-Facturation-Simulation-de-cr-dit-VBE-EXCEL/assets/143803507/ff987951-47ec-4fd1-9d25-38b0a84f392a)


## Description du Projet

Ce projet a pour but de développer une application en VBA Excel pour automatiser la facturation et fournir un simulateur de crédit pour une entreprise de vente d’équipements informatiques. L'application réduit les erreurs de calcul et améliore l'efficacité du processus de facturation.

## Table des Matières

- [Fonctionnalités](#fonctionnalités)
- [Installation](#installation)
- [Utilisation](#utilisation)
- [Structure du Code](#structure-du-code)
- [Contributions](#contributions)
- [Licence](#licence)

## Fonctionnalités

L'application comporte les fonctionnalités suivantes :

1. **Automatisation de la Facturation**
    - Saisie du numéro du client pour afficher les informations de la facture.
    - Saisie des références des produits pour afficher les désignations et les prix unitaires.
    - Calcul des montants hors taxes (HT), des remises, des frais de transport, de la TVA, et du montant total TTC.

2. **Gestion des Clients**
    - Ajout de nouveaux clients via un formulaire utilisateur (UserForm).

3. **Gestion des Produits**
    - Ajout de nouveaux produits au catalogue via un formulaire utilisateur (UserForm).

4. **Simulation de Crédit**
    - Simulation d'emprunt basée sur les montants de la facture, avec des taux annuels et des durées personnalisables.



## Installation

1. **Cloner le dépôt**
    ```bash
    git clone https://github.com/HAMZA25032005/Automatisation-de-Facturation-Simulation-de-cr-dit-VBE-EXCEL.git
    ```
2. **Ouvrir le fichier Excel**
    - Ouvrez le fichier Excel fourni dans le dépôt avec Microsoft Excel.

3. **Activer les macros**
    - Assurez-vous que les macros sont activées dans Excel pour permettre l'exécution du code VBA.

## Utilisation

### Automatisation de la Facturation

1. **Remplir les informations de la facture**
    - Saisissez le numéro du client pour afficher la raison sociale et l’adresse.
    - Saisissez les références des produits pour afficher les désignations et les prix unitaires.
    - Renseignez manuellement les quantités.

2. **Calculer les montants**
    - Le programme calculera automatiquement les montants HT, les remises, les frais de transport, la TVA, et le montant TTC.

### Gestion des Clients

1. **Ajouter un client**
    - Ouvrez le formulaire utilisateur pour saisir les informations du nouveau client.

### Gestion des Produits

1. **Ajouter un produit**
    - Ouvrez le formulaire utilisateur pour saisir les informations du nouveau produit.

### Simulation de Crédit

1. **Simuler un emprunt**
    - Saisissez le taux annuel et la durée en mois pour générer le tableau de simulation.

## Structure du Code

- **Module Facturation**
    - Contient les fonctions pour calculer les montants HT, les remises, les frais de transport, la TVA, et le montant TTC.

- **Module Clients**
    - Contient les fonctions pour ajouter et gérer les clients.

- **Module Produits**
    - Contient les fonctions pour ajouter et gérer les produits.

- **Module Simulation de Crédit**
    - Contient les fonctions pour simuler les emprunts.

## Contributions

Les contributions sont les bienvenues ! Veuillez suivre les étapes suivantes pour contribuer :

1. Fork le projet
2. Créez votre branche feature (`git checkout -b feature/ma-nouvelle-fonctionnalité`)
3. Commitez vos modifications (`git commit -m 'Ajouter ma nouvelle fonctionnalité'`)
4. Poussez votre branche (`git push origin feature/ma-nouvelle-fonctionnalité`)
5. Ouvrez une Pull Request

## Licence

Ce projet est sous licence MIT. Voir le fichier `LICENSE` pour plus de détails.
