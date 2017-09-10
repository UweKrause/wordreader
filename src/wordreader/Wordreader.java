package wordreader;

import org.apache.poi.poifs.filesystem.*;
import org.apache.poi.hwpf.*;
import org.apache.poi.hwpf.extractor.*;
import java.io.*;
import java.util.Arrays;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Wordreader {

    /*
     * Eine Word-Datei wird eingelesen und die in ihr vorkommenden Worte werden
     * # alphabetisch sortiert mit Angabe der Haeufigkeit
     * # nach Haufigkeit sortiert
     */
    public static void main(String[] args) {

        /**
         * Erstelle eine Liste mit allen in ./IN/ abgelegten Word-Dateien
         */
        //String rootpfad = "/home/uwe/NetBeansProjects/Wordreader/";
        //System.err.println(System.getProperty("user.dir"));
        String rootpfad = System.getProperty("user.dir");

        File[] fileArray = new File(rootpfad + "/IN/").listFiles();

        List<String> liste = Arrays.stream(fileArray)
                .map(File::getName)
                .filter(e -> e.endsWith(".doc"))
                // oder fuer mehrere moegliche Endungen? .docx ?
                //.filter(e -> e.endsWith(".doc") || e.endsWith(".docx"))

                .collect(Collectors.toList());

        System.out.println("Bearbeite folgende Dateien:");
        System.out.println(liste);
        System.out.println("");

        /**
         * bearbeite alle gefundenen Dateien
         */
        for (String filename : liste) {

            POIFSFileSystem fs = null;

            /**
             * Ergebnisdateien anlegen
             */
            try {

                /**
                 * Bereite Ergebnisdatei fuer alphabetische Sortierung vor
                 */
                File alpha = new File(rootpfad + "/OUT/" + filename + "_alphabetisch.txt");
                // wenn Datei nicht existiert, wird sie neu angelegt
                if (!alpha.exists()) {
                    alpha.createNewFile();
                }
                FileWriter alpha_fw = new FileWriter(alpha.getAbsoluteFile());
                BufferedWriter alpha_bw = new BufferedWriter(alpha_fw);
                StringBuilder alpha_sb = new StringBuilder();

                /**
                 * Bereite Ergebnisdatei fuer numerische Sortierung vor
                 */
                File numeric = new File(rootpfad + "/OUT/" + filename + "_numerisch.txt");
                // wenn Datei nicht existiert, wird sie neu angelegt
                if (!numeric.exists()) {
                    numeric.createNewFile();
                }
                FileWriter numeric_fw = new FileWriter(numeric.getAbsoluteFile());
                BufferedWriter numeric_bw = new BufferedWriter(numeric_fw);
                StringBuilder numeric_sb = new StringBuilder();

                /**
                 * lese Quelldateien ein
                 */
                try {

                    fs = new POIFSFileSystem(new FileInputStream("IN/" + filename));
                    HWPFDocument doc = new HWPFDocument(fs);
                    WordExtractor we = new WordExtractor(doc);

                    /**
                     * Durchlaeuft Datei, packt jedes Wort in eine Map mit
                     * Haeufigkeit des Vorkommens
                     */
                    Map<String, Integer> wc = new TreeMap<>();

                    String[] text = we.getParagraphText();

                    Arrays.stream(text).parallel()
                            // nur Zeilen mit Inhalt
                            .filter(x -> x.length() > 0)
                            // liefert einen Stream von Worten, getrennt am Leerzeichen
                            .flatMap(txt -> Stream.of(txt.split(" ")))
                            // entfernt unerwuenschte Zeichen
                            .map(word -> word.replaceAll("[^1-9a-zA-ZäöüÄÖÜß-]", ""))
                            // Filtert Fragmente, bei denen nichts mehr uebrig geblieben ist
                            .filter(x -> x.length() > 0)
                            // macht jedes Wort zu lowercase, um worte am Satzanfang nicht anders zu zaehlen als zwischendurch
                            .map(word -> word.toLowerCase())
                            // schmeisst die Worte in die Map
                            // wenn schon drin index + 1, ansonsten Value = 1
                            .forEachOrdered((String word) -> {
                                if (wc.containsKey(word)) {
                                    wc.put(word, wc.get(word) + 1);
                                } else {
                                    wc.put(word, 1);
                                }
                            });

                    /**
                     * Erste
                     */
                    Map<Integer, List<String>> mol = new TreeMap<>();

                    wc.entrySet().stream().forEach((entry) -> {
                        String string = entry.getKey();
                        Integer anzahl = entry.getValue();

                        mol.putIfAbsent(anzahl, new LinkedList<>());

                        mol.get(anzahl).add(string);
                    });

                    /**
                     * sortiert nach Alphabet
                     */
                    wc.entrySet().stream()
                            .forEach(entry -> alpha_sb.append(entry.getKey()).append(" ").append(entry.getValue()).append(("\r\n")));

                    /**
                     * Sortiert nach Anzahl
                     */
                    mol.entrySet().stream()
                            .map((entry) -> {
                                Integer anzahl = entry.getKey();
                                List<String> worte = entry.getValue();
                                /*
                                        System.out.println(anzahl);
                                        System.out.println(worte);
                                 */
                                numeric_sb.append(anzahl).append("\r\n");
                                numeric_sb.append(worte).append("\r\n");
                                return entry;
                            }).forEach((_item) -> {
                        numeric_sb.append("\r\n");
                        //System.out.println("");
                    });

                } catch (IOException e) {
                    e.printStackTrace();
                }

                alpha_bw.write(alpha_sb.toString());
                alpha_bw.close();

                numeric_bw.write(numeric_sb.toString());
                numeric_bw.close();

            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
