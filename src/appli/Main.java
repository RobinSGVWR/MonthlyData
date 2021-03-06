package appli;

import javax.swing.*;

public class Main {
    public static void main(String [ ] args)
    {
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        }catch (IllegalAccessException | InstantiationException | UnsupportedLookAndFeelException | ClassNotFoundException e) {
            e.printStackTrace();
        }

        //GESTION FENETRE

        JFrame fenetre = new JFrame();
        fenetre.setResizable(false);


        ExcelReader reader = new ExcelReader(fenetre);
        //Définit un titre pour notre fenêtre
        fenetre.setTitle("remplissage Excel");
        //Termine le processus lorsqu'on clique sur la croix rouge
        fenetre.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        //Et enfin, la rendre visible
        fenetre.setVisible(true);
        //Instanciation d'un objet JPanel

        //Définition de sa couleur de fond
        //On prévient notre JFrame que notre JPanel sera son content pane
        fenetre.setContentPane(reader);
        fenetre.pack();
        fenetre.setVisible(true);





    }

}
