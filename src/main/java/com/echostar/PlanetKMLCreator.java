package com.echostar;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Vector;
import java.util.concurrent.ExecutionException;
import java.util.stream.Collectors;

/**
 * A Java Swing application that allows a user to select a large Excel file and view
 * specific sheets in a JTable using a memory-efficient SAX parser to prevent OutOfMemoryError.
 * This version opens in full-screen, preserves column order, and performs advanced data merging for the Sectors tab.
 * It also includes a progress bar for long-running operations.
 */
public class PlanetKMLCreator extends JFrame {

    private final JTabbedPane tabbedPane;
    private final JLabel statusLabel;
    private final JProgressBar progressBar;
    private final JButton kmlButton;
    private DefaultTableModel controllersModel;
    private SheetData finalSectorsData;
    private SheetData finalSiteData;

    /**
     * Helper class to hold both the ordered headers and the data from a sheet.
     */
    private static class SheetData {
        final List<String> headers;
        final List<Map<String, String>> tableData;

        SheetData(List<String> headers, List<Map<String, String>> tableData) {
            this.headers = headers;
            this.tableData = tableData;
        }
    }

    // Helper class to store user's KML choices for each band
    private static class BandSettings {
        boolean include = true;
        Color color = Color.BLUE;
        int size = 500; // Default size in meters
        int transparency = 50; // Default transparency (0-100)
    }

    public PlanetKMLCreator() {
        super("KML Generator V1.2");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setExtendedState(JFrame.MAXIMIZED_BOTH);
        setSize(1280, 800);
        setLocationRelativeTo(null);

        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel topPanel = new JPanel(new BorderLayout(10, 10));
        JButton openButton = new JButton("Open Planet Export");
        kmlButton = new JButton("Generate KML");
        kmlButton.setEnabled(false); // Disabled by default
        statusLabel = new JLabel("No file selected. Please open a large .xlsx file.");
        
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        buttonPanel.add(openButton);
        buttonPanel.add(kmlButton);
        
        topPanel.add(buttonPanel, BorderLayout.WEST);
        topPanel.add(statusLabel, BorderLayout.CENTER);

        tabbedPane = new JTabbedPane();
        mainPanel.add(topPanel, BorderLayout.NORTH);
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        
        // --- BOTTOM PANEL WITH PROGRESS BAR ---
        JPanel bottomPanel = new JPanel();
        bottomPanel.setLayout(new BoxLayout(bottomPanel, BoxLayout.Y_AXIS));
        
        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setVisible(false); // Initially hidden
        
        JLabel supportLabel = new JLabel("<html><i><b><font color='red'>For Support Email: jayeshkumar.patel@dish.com</font></b></i></html>");
        supportLabel.setAlignmentX(Component.CENTER_ALIGNMENT);

        bottomPanel.add(progressBar);
        bottomPanel.add(Box.createRigidArea(new Dimension(0, 5))); // Spacer
        bottomPanel.add(supportLabel);
        mainPanel.add(bottomPanel, BorderLayout.SOUTH);
        // --- END BOTTOM PANEL ---

        add(mainPanel);

        tabbedPane.addTab("Controllers", createControllersPanel());

        openButton.addActionListener(e -> openFile());
        kmlButton.addActionListener(e -> generateKML());
    }

    private JPanel createControllersPanel() {
        JPanel controllerPanel = new JPanel(new BorderLayout(5, 5));
        String[] columnNames = {"Electrical Controller", "Band"};
        Object[][] data = {
            {"R1", "LB Electrical Tilt"}, {"R2", "LB Electrical Tilt"}, {"B", "MB Electrical Tilt"},
            {"Controller_617-894_12", "LB Electrical Tilt"}, {"Controller_617-894_34", "LB Electrical Tilt"},
            {"Controller 1", "LB Electrical Tilt"}, {"Controller_1695-2690_56", "MB Electrical Tilt"},
            {"Controller_1695-2690_78", "MB Electrical Tilt"}, {"Controller 2", "MB Electrical Tilt"},
            {"Controller 3", "MB Electrical Tilt"}, {"Port 1-2", "LB Electrical Tilt"},
            {"Y1", "MB Electrical Tilt"}, {"Y2", "MB Electrical Tilt"}, {"Port 3-4", "MB Electrical Tilt"},
            {"Port 5-6", "MB Electrical Tilt"}, {"Port 1-4", "LB Electrical Tilt"},
            {"Port 5-8", "MB Electrical Tilt"}, {"Port 9-10", "MB Electrical Tilt"},
            {"Port 3-4", "LB Electrical Tilt"}, {"Port 7-8", "MB Electrical Tilt"},
            {"R1 LB Controller", "LB Electrical Tilt"}, {"R2 LB Controller", "LB Electrical Tilt"},
            {"Y1 HB Controller", "MB Electrical Tilt"}, {"Y2 HB Controller", "MB Electrical Tilt"}
        };

        controllersModel = new DefaultTableModel(data, columnNames);
        JTable table = new JTable(controllersModel);
        table.setFillsViewportHeight(true);
        
        JButton addRowButton = new JButton("Add Row");
        addRowButton.addActionListener(e -> controllersModel.addRow(new Object[]{"", ""}));
        
        JPanel buttonPanelSouth = new JPanel(new FlowLayout(FlowLayout.CENTER));
        buttonPanelSouth.add(addRowButton);
        
        controllerPanel.add(new JScrollPane(table), BorderLayout.CENTER);
        controllerPanel.add(buttonPanelSouth, BorderLayout.SOUTH);
        return controllerPanel;
    }

    private void openFile() {
        // Confirmation Dialog
        String confirmationMessage = "Do you have the latest Planet sheets required for processing?\n" +
                                     "(Antennas, Antenna_Electrical_Parameters, Sectors, NR_Sector_Carriers, Sites)";
        int response = JOptionPane.showConfirmDialog(this, confirmationMessage, "Confirm Sheets", JOptionPane.YES_NO_OPTION, JOptionPane.QUESTION_MESSAGE);

        if (response == JOptionPane.YES_OPTION) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Select an Excel File");
            fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
                public boolean accept(File f) {
                    return f.getName().toLowerCase().endsWith(".xlsx") || f.isDirectory();
                }
                public String getDescription() {
                    return "Excel Files (*.xlsx)";
                }
            });

            if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                statusLabel.setText("Loading file: " + selectedFile.getName());
                // Use SwingWorker to process the file in the background
                ExcelLoaderTask task = new ExcelLoaderTask(selectedFile);
                task.execute();
            }
        } else {
            JOptionPane.showMessageDialog(this, "Please come back with the necessary Planet export.", "Process Canceled", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    /**
     * SwingWorker to load and process the Excel file in the background,
     * updating the GUI with progress.
     */
    private class ExcelLoaderTask extends SwingWorker<Map<String, SheetData>, String> {
        private final File excelFile;

        ExcelLoaderTask(File excelFile) {
            this.excelFile = excelFile;
        }

        @Override
        protected void process(List<String> chunks) {
            // Update status label with the latest message from publish()
            statusLabel.setText(chunks.get(chunks.size() - 1));
            progressBar.setValue(progressBar.getValue() + 1);
        }

        @Override
        protected Map<String, SheetData> doInBackground() throws Exception {
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(0);
                progressBar.setMaximum(8); // 5 sheets to read + 3 processing steps
                progressBar.setVisible(true);
                // Clear old tabs
                for (int i = tabbedPane.getTabCount() - 1; i >= 0; i--) {
                    if (!tabbedPane.getTitleAt(i).equals("Controllers")) {
                        tabbedPane.remove(i);
                    }
                }
            });

            String[] sheetsToRead = { "Antennas", "Antenna_Electrical_Parameters", "Sectors", "NR_Sector_Carriers", "Sites" };
            Map<String, SheetData> allSheetsData = new HashMap<>();

            for (String sheetName : sheetsToRead) {
                publish("Processing sheet: " + sheetName + "...");
                SheetData sheetData = processSheetWithSAX(excelFile, sheetName);
                if (sheetData != null) {
                    allSheetsData.put(sheetName, sheetData);
                }
            }

            publish("Processing Electrical Parameters...");
            SheetData processedElectricalParams = processElectricalParametersData(allSheetsData.get("Antenna_Electrical_Parameters"));
            allSheetsData.put("Processed_Electrical_Parameters", processedElectricalParams);

            publish("Processing Site Data...");
            finalSiteData = processSiteData(allSheetsData.get("Sites"), allSheetsData.get("Antennas"));

            publish("Processing Sectors Data...");
            finalSectorsData = processSectorsData(
                allSheetsData.get("Sectors"), allSheetsData.get("NR_Sector_Carriers"),
                allSheetsData.get("Antennas"), processedElectricalParams
            );
            
            return allSheetsData;
        }

        @Override
        protected void done() {
            try {
                get(); // Call get() to rethrow any exception that occurred during doInBackground
                
                // Display the Site tab
                if (finalSiteData != null && !finalSiteData.tableData.isEmpty()) {
                    JTable table = new JTable(createTableModel(finalSiteData));
                    table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
                    tabbedPane.addTab("Sites", new JScrollPane(table));
                }

                // Display the final Sectors tab
                if (finalSectorsData != null && !finalSectorsData.tableData.isEmpty()) {
                    JTable table = new JTable(createTableModel(finalSectorsData));
                    table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
                    tabbedPane.addTab("Sectors", new JScrollPane(table));
                }
                
                statusLabel.setText("Successfully loaded and processed: " + excelFile.getName());
                kmlButton.setEnabled(true); // Enable KML button on success

            } catch (InterruptedException | ExecutionException e) {
                statusLabel.setText("Error processing file: " + e.getCause().getMessage());
                e.printStackTrace();
                JOptionPane.showMessageDialog(PlanetKMLCreator.this, "Failed to process Excel file: \n" + e.getCause().getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                kmlButton.setEnabled(false); // Keep it disabled on error
            } finally {
                progressBar.setVisible(false);
            }
        }
    }


    private SheetData processSiteData(SheetData siteData, SheetData antennasData) {
        if (siteData == null) return null;
        Map<String, String> heightLookup = new HashMap<>();
        if (antennasData != null) {
            for (Map<String, String> antennaRow : antennasData.tableData) {
                String siteId = antennaRow.getOrDefault("Site ID", "");
                if (!siteId.isEmpty() && !heightLookup.containsKey(siteId)) {
                    heightLookup.put(siteId, antennaRow.getOrDefault("Height (ft)", ""));
                }
            }
        }
        List<String> finalHeaders = Arrays.asList("Site ID", "Longitude", "Latitude", "Site Name", "Custom: Cluster_ID", "Custom: gNodeB_Id", "Custom: gNodeB_Site_Number", "Custom: TAC", "Height (ft)");
        List<Map<String, String>> processedData = new ArrayList<>();
        for (Map<String, String> row : siteData.tableData) {
            Map<String, String> newRow = new LinkedHashMap<>();
            String siteId = row.getOrDefault("Site ID", "");
            for (String header : finalHeaders) {
                if (header.equals("Height (ft)")) {
                    newRow.put(header, heightLookup.getOrDefault(siteId, ""));
                } else {
                    newRow.put(header, row.getOrDefault(header, ""));
                }
            }
            processedData.add(newRow);
        }
        return new SheetData(finalHeaders, processedData);
    }

    private SheetData processSectorsData(SheetData sectors, SheetData nrCarriers, SheetData antennas, SheetData electricalParams) {
        if (sectors == null || nrCarriers == null || antennas == null || electricalParams == null) return null;
        Map<String, String> pciLookup = new HashMap<>();
        for (Map<String, String> row : nrCarriers.tableData) {
            pciLookup.put(row.getOrDefault("Site ID", "") + "||" + row.getOrDefault("Sector ID", ""), row.get("Physical Cell ID"));
        }
        Map<String, Map<String, String>> antennaLookup = new HashMap<>();
        for (Map<String, String> row : antennas.tableData) {
            antennaLookup.put(row.getOrDefault("Site ID", "") + "||" + row.getOrDefault("Antenna ID", ""), row);
        }
        Map<String, String> electricalTiltLookup = new HashMap<>();
        for (Map<String, String> paramsRow : electricalParams.tableData) {
            String key = paramsRow.getOrDefault("Site ID", "") + paramsRow.getOrDefault("Antenna ID", "") + paramsRow.getOrDefault("Band Info", "");
            if (!key.isEmpty()) {
                electricalTiltLookup.put(key, paramsRow.getOrDefault("Electrical Tilt", ""));
            }
        }
        List<String> finalHeaders = Arrays.asList("Site ID", "Band Name", "Custom: NR_Cell_Global_ID", "Custom: NR_Cell_Name", "Custom: RU_Model", "Sector ID", "Physical Cell ID", "Antenna ID", "Latitude", "Longitude", "Antenna File", "Height (ft)", "Azimuth", "Electrical Tilt");
        List<Map<String, String>> processedData = new ArrayList<>();
        for (Map<String, String> sectorRow : sectors.tableData) {
            Map<String, String> newRow = new LinkedHashMap<>();
            String siteId = sectorRow.getOrDefault("Site ID", "");
            String originalSectorId = sectorRow.getOrDefault("Sector ID", "");
            if (originalSectorId.isEmpty() || siteId.isEmpty()) continue;
            String antennaId = "";
            if (!originalSectorId.isEmpty()) {
                char lastChar = originalSectorId.charAt(originalSectorId.length() - 1);
                if (Character.isDigit(lastChar)) antennaId = String.valueOf(lastChar);
            }
            String bandName = sectorRow.getOrDefault("Band Name", "").toUpperCase();
            String bandInfo = (bandName.startsWith("N29") || bandName.startsWith("N71")) ? "LB Electrical Tilt" : "MB Electrical Tilt";
            String key = siteId + antennaId + bandInfo;
            String electricalTilt = electricalTiltLookup.getOrDefault(key, "");
            Map<String, String> antennaData = antennaLookup.get(siteId + "||" + antennaId);
            newRow.put("Site ID", siteId);
            newRow.put("Band Name", sectorRow.getOrDefault("Band Name", ""));
            newRow.put("Custom: NR_Cell_Global_ID", sectorRow.getOrDefault("Custom: NR_Cell_Global_Id", ""));
            newRow.put("Custom: NR_Cell_Name", sectorRow.getOrDefault("Custom: NR_Cell_Name", ""));
            newRow.put("Custom: RU_Model", sectorRow.getOrDefault("Custom: RU_Model", ""));
            newRow.put("Sector ID", originalSectorId);
            newRow.put("Physical Cell ID", pciLookup.getOrDefault(siteId + "||" + originalSectorId, ""));
            newRow.put("Antenna ID", antennaId);
            if (antennaData != null) {
                newRow.put("Latitude", antennaData.getOrDefault("Latitude", ""));
                newRow.put("Longitude", antennaData.getOrDefault("Longitude", ""));
                newRow.put("Antenna File", antennaData.getOrDefault("Antenna File", "").replace(".pafx", ""));
                newRow.put("Height (ft)", antennaData.getOrDefault("Height (ft)", ""));
                newRow.put("Azimuth", antennaData.getOrDefault("Azimuth", ""));
            } else {
                newRow.put("Latitude", ""); newRow.put("Longitude", ""); newRow.put("Antenna File", "");
                newRow.put("Height (ft)", ""); newRow.put("Azimuth", "");
            }
            newRow.put("Electrical Tilt", electricalTilt);
            processedData.add(newRow);
        }
        return new SheetData(finalHeaders, processedData);
    }

    private SheetData processElectricalParametersData(SheetData electricalParams) {
        if (electricalParams == null) return null;
        Map<String, String> controllerBandLookup = new HashMap<>();
        for (int i = 0; i < controllersModel.getRowCount(); i++) {
            String controller = (String) controllersModel.getValueAt(i, 0);
            String band = (String) controllersModel.getValueAt(i, 1);
            if (controller != null && !controller.isEmpty()) {
                controllerBandLookup.put(controller, band);
            }
        }
        List<String> finalHeaders = new ArrayList<>(electricalParams.headers);
        if (!finalHeaders.contains("Band Info")) finalHeaders.add("Band Info");
        List<Map<String, String>> processedData = new ArrayList<>();
        for (Map<String, String> electricalRow : electricalParams.tableData) {
            Map<String, String> newRow = new LinkedHashMap<>(electricalRow);
            String controller = electricalRow.getOrDefault("Electrical Controller", "");
            newRow.put("Band Info", controllerBandLookup.getOrDefault(controller, ""));
            processedData.add(newRow);
        }
        return new SheetData(finalHeaders, processedData);
    }
    
    private void generateKML() {
        if (finalSectorsData == null || finalSectorsData.tableData.isEmpty() || finalSiteData == null || finalSiteData.tableData.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No data in the Sectors or Site tab to generate KML.", "No Data", JOptionPane.WARNING_MESSAGE);
            return;
        }

        Set<String> uniqueBands = finalSectorsData.tableData.stream().map(row -> row.getOrDefault("Band Name", "Unknown")).collect(Collectors.toCollection(LinkedHashSet::new));
        Map<String, BandSettings> bandSettings = showBandCustomizationDialog(uniqueBands);
        if (bandSettings == null) return;

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save KML File");
        String firstSiteId = finalSiteData.tableData.get(0).getOrDefault("Site ID", "SITE");
        String sitePart = (firstSiteId.length() >= 5) ? firstSiteId.substring(2, 5) : "SITE";
        String datePart = new SimpleDateFormat("MMddyyyy").format(new Date());
        fileChooser.setSelectedFile(new File(sitePart + "_" + datePart + ".kml"));
        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            public boolean accept(File f) { return f.getName().toLowerCase().endsWith(".kml") || f.isDirectory(); }
            public String getDescription() { return "KML Files (*.kml)"; }
        });

        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            KMLGeneratorTask task = new KMLGeneratorTask(fileToSave, bandSettings, uniqueBands);
            task.execute();
        }
    }

    /**
     * SwingWorker to generate the KML file in the background.
     */
    private class KMLGeneratorTask extends SwingWorker<Void, Integer> {
        private final File fileToSave;
        private final Map<String, BandSettings> bandSettings;
        private final Set<String> uniqueBands;

        KMLGeneratorTask(File fileToSave, Map<String, BandSettings> bandSettings, Set<String> uniqueBands) {
            this.fileToSave = fileToSave;
            this.bandSettings = bandSettings;
            this.uniqueBands = uniqueBands;
        }

        @Override
        protected void process(List<Integer> chunks) {
            progressBar.setValue(chunks.get(chunks.size() - 1));
        }

        @Override
        protected Void doInBackground() throws Exception {
            int totalPlacemarks = finalSiteData.tableData.size() + finalSectorsData.tableData.size() * 3; // Approximation for sectors + labels
            SwingUtilities.invokeLater(() -> {
                progressBar.setValue(0);
                progressBar.setMaximum(totalPlacemarks);
                progressBar.setVisible(true);
                statusLabel.setText("Generating KML file...");
            });

            int progress = 0;
            try (FileWriter writer = new FileWriter(fileToSave)) {
                writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<kml xmlns=\"http://www.opengis.net/kml/2.2\">\n<Document>\n");
                writer.write(getSiteStyle());
                writer.write("<Style id=\"label-style\"><IconStyle><scale>0</scale></IconStyle><LabelStyle><color>ffffffff</color><scale>0.8</scale></LabelStyle></Style>\n");
                for (Map.Entry<String, BandSettings> entry : bandSettings.entrySet()) {
                    if (entry.getValue().include) {
                        writer.write(createKMLStyle(entry.getKey(), entry.getValue().color, entry.getValue().transparency));
                    }
                }

                writer.write("<Folder>\n<name>SITES</name>\n");
                for (Map<String, String> siteRow : finalSiteData.tableData) {
                    writer.write(createSitePlacemark(siteRow));
                    publish(++progress);
                }
                writer.write("</Folder>\n");

                writer.write("<Folder>\n<name>SECTORS</name>\n");
                Map<String, List<Map<String, String>>> sectorsByBand = finalSectorsData.tableData.stream().collect(Collectors.groupingBy(row -> row.getOrDefault("Band Name", "Unknown")));
                
                List<String> bandOrder = new ArrayList<>(uniqueBands);
                bandOrder.sort((band1, band2) -> {
                    BandSettings settings1 = bandSettings.get(band1);
                    BandSettings settings2 = bandSettings.get(band2);
                    return Integer.compare(settings2.size, settings1.size);
                });

                for (int i = 0; i < bandOrder.size(); i++) {
                    String bandName = bandOrder.get(i);
                    List<Map<String, String>> rowsForBand = sectorsByBand.get(bandName);
                    BandSettings settings = bandSettings.get(bandName);
                    if (settings != null && settings.include && rowsForBand != null) {
                        writer.write("<Folder>\n<name>" + bandName + "</name>\n");
                        for (Map<String, String> row : rowsForBand) {
                            writer.write(createSectorPlacemark(row, bandName, settings.size, i));
                            publish(++progress);
                        }
                        writer.write("</Folder>\n");
                    }
                }
                writer.write("</Folder>\n");

                writer.write("<Folder>\n<name>Display</name>\n");
                List<String> displayHeaders = Arrays.asList("Physical Cell ID", "Electrical Tilt", "Azimuth");
                for (String header : displayHeaders) {
                    writer.write("<Folder>\n<name>" + header + "</name>\n");
                    for (Map.Entry<String, List<Map<String, String>>> bandEntry : sectorsByBand.entrySet()) {
                        String bandName = bandEntry.getKey();
                        BandSettings settings = bandSettings.get(bandName);
                        boolean createBandFolder = header.equals("Electrical Tilt") || bandName.toUpperCase().contains("N71");
                        if (createBandFolder && settings != null && settings.include) {
                            writer.write("<Folder>\n<name>" + bandName + "</name>\n");
                            for (Map<String, String> row : bandEntry.getValue()) {
                                writer.write(createLabelPlacemark(row, header, settings.size));
                                publish(++progress);
                            }
                            writer.write("</Folder>\n");
                        }
                    }
                    writer.write("</Folder>\n");
                }
                writer.write("</Folder>\n");

                writer.write("</Document>\n</kml>\n");
            }
            return null;
        }

        @Override
        protected void done() {
            try {
                get();
                JOptionPane.showMessageDialog(PlanetKMLCreator.this, "KML file generated successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
                statusLabel.setText("KML file saved to " + fileToSave.getName());
            } catch (InterruptedException | ExecutionException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(PlanetKMLCreator.this, "Error generating KML file: " + e.getCause().getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
                statusLabel.setText("Error generating KML file.");
            } finally {
                progressBar.setVisible(false);
            }
        }
    }

    private Map<String, BandSettings> showBandCustomizationDialog(Set<String> bands) {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(2, 5, 2, 5);
        gbc.gridx = 0; gbc.gridy = 0;

        Map<String, JCheckBox> checkBoxes = new HashMap<>();
        Map<String, JButton> colorButtons = new HashMap<>();
        Map<String, JSpinner> sizeSpinners = new HashMap<>();
        Map<String, JSlider> transparencySliders = new HashMap<>();
        Map<String, JSpinner> transparencySpinners = new HashMap<>();
        Map<String, BandSettings> settingsMap = new HashMap<>();

        panel.add(new JLabel("Include"), gbc); gbc.gridx++;
        panel.add(new JLabel("Band Name"), gbc); gbc.gridx++;
        panel.add(new JLabel("Color"), gbc); gbc.gridx++;
        panel.add(new JLabel("Size (m)"), gbc); gbc.gridx++;
        gbc.fill = GridBagConstraints.HORIZONTAL;
        panel.add(new JLabel("Transparency (%)"), gbc); gbc.fill = GridBagConstraints.NONE;
        gbc.gridx++;
        panel.add(new JLabel(""), gbc); // Empty header for the percentage value
        gbc.gridy++;
        
        Color[] brightColors = {Color.CYAN, Color.MAGENTA, Color.YELLOW, Color.GREEN, Color.ORANGE, Color.PINK};
        int colorIndex = 0;

        for (String band : bands) {
            BandSettings settings = new BandSettings();
            String upperBand = band.toUpperCase();
            if (upperBand.contains("N71")) settings.size = 500;
            else if (upperBand.contains("N70")) settings.size = 400;
            else if (upperBand.contains("N66")) settings.size = 350;
            else if (upperBand.contains("N29")) settings.size = 300;
            else settings.size = 500;
            settings.color = brightColors[colorIndex++ % brightColors.length];
            settingsMap.put(band, settings);

            gbc.gridx = 0;
            JCheckBox checkBox = new JCheckBox("", true);
            checkBoxes.put(band, checkBox);
            panel.add(checkBox, gbc);

            gbc.gridx++; panel.add(new JLabel(band), gbc);

            gbc.gridx++;
            JButton colorButton = new JButton(" ");
            colorButton.setBackground(settings.color);
            colorButton.addActionListener(e -> {
                Color newColor = JColorChooser.showDialog(null, "Choose a color", colorButton.getBackground());
                if (newColor != null) {
                    colorButton.setBackground(newColor);
                    settingsMap.get(band).color = newColor;
                }
            });
            colorButtons.put(band, colorButton);
            panel.add(colorButton, gbc);

            gbc.gridx++;
            JSpinner sizeSpinner = new JSpinner(new SpinnerNumberModel(settings.size, 1, 10000, 100));
            sizeSpinners.put(band, sizeSpinner);
            panel.add(sizeSpinner, gbc);
            
            gbc.gridx++; gbc.fill = GridBagConstraints.HORIZONTAL;
            JSlider transparencySlider = new JSlider(0, 100, settings.transparency);
            transparencySlider.setMajorTickSpacing(5);
            transparencySlider.setSnapToTicks(true);
            transparencySliders.put(band, transparencySlider);
            panel.add(transparencySlider, gbc); gbc.fill = GridBagConstraints.NONE;

            gbc.gridx++;
            SpinnerModel spinnerModel = new SpinnerNumberModel(settings.transparency, 0, 100, 5);
            JSpinner transparencySpinner = new JSpinner(spinnerModel);
            transparencySpinner.setPreferredSize(new Dimension(60, 25));
            transparencySpinners.put(band, transparencySpinner);
            panel.add(transparencySpinner, gbc);

            transparencySlider.addChangeListener(e -> {
                JSlider source = (JSlider) e.getSource();
                transparencySpinner.setValue(source.getValue());
            });
            
            transparencySpinner.addChangeListener(e -> {
                JSpinner source = (JSpinner) e.getSource();
                transparencySlider.setValue((Integer) source.getValue());
            });
            
            gbc.gridy++;
        }
        
        JScrollPane scrollPane = new JScrollPane(panel);
        scrollPane.setPreferredSize(new Dimension(650, 400)); // Widen the dialog slightly

        int result = JOptionPane.showConfirmDialog(this, scrollPane, "Customize KML Bands", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (result == JOptionPane.OK_OPTION) {
            for (String band : bands) {
                BandSettings settings = settingsMap.get(band);
                settings.include = checkBoxes.get(band).isSelected();
                settings.size = (int) sizeSpinners.get(band).getValue();
                settings.transparency = (int) transparencySpinners.get(band).getValue();
            }
            return settingsMap;
        }
        return null;
    }

    private String getSiteStyle() {
        // Boost Mobile orange: #f26522. KML color (AABBGGRR): ff2265f2
        String boostOrange = "ff2265f2";
        return "<Style id=\"normPointStyle\">\n" +
               "    <IconStyle>\n" +
               "        <scale>0.8</scale>\n" +
               "        <Icon>\n" +
               "            <href>https://i.ibb.co/5YtdGtG/LOGO-PLOT-TRNS.png</href>\n" +
               "        </Icon>\n" +
               "        <hotSpot x=\"0.5\" y=\"0.5\" xunits=\"fraction\" yunits=\"fraction\"/>\n" +
               "    </IconStyle>\n" +
               "    <LabelStyle>\n" +
               "        <color>ffffffff</color>\n" +
               "        <scale>1.0</scale>\n" +
               "    </LabelStyle>\n" +
               "    <LineStyle>\n" +
               "        <color>" + boostOrange + "</color>\n" +
               "        <width>10</width>\n" +
               "    </LineStyle>\n" +
               "</Style>\n" +
               "<StyleMap id=\"site-icon\">\n" +
               "    <Pair>\n" +
               "        <key>normal</key>\n" +
               "        <styleUrl>#normPointStyle</styleUrl>\n" +
               "    </Pair>\n" +
               "    <Pair>\n" +
               "        <key>highlight</key>\n" +
               "        <styleUrl>#normPointStyle</styleUrl>\n" +
               "    </Pair>\n" +
               "</StyleMap>\n" +
               "<Style id=\"site-line\"><LineStyle><color>" + boostOrange + "</color><width>15</width></LineStyle></Style>\n";
    }

    private String createKMLStyle(String id, Color color, int transparencyPercent) {
        String safeId = id.replaceAll("[^a-zA-Z0-9]", "");
        
        // Calculate opacity from the transparency percentage.
        // In KML, alpha FF is opaque, 00 is transparent.
        // The slider is for transparency (0=opaque, 100=transparent), so we convert it to opacity for KML.
        int opacityPercent = 100 - transparencyPercent;
        int alpha = (int) Math.round(opacityPercent * 2.55); // (percent/100) * 255
        String alphaHex = String.format("%02x", alpha);
        
        // Full KML color string (AABBGGRR)
        String kmlColor = String.format("%s%02x%02x%02x", alphaHex, color.getBlue(), color.getGreen(), color.getRed());
        
        return String.format("<Style id=\"%s\"><LineStyle><color>ff%s</color></LineStyle><PolyStyle><color>%s</color></PolyStyle></Style>\n", safeId, kmlColor.substring(2), kmlColor);
    }

    private String createSitePlacemark(Map<String, String> row) {
        String siteId = row.getOrDefault("Site ID", "N/A");
        String lon = row.getOrDefault("Longitude", "0");
        String lat = row.getOrDefault("Latitude", "0");
        String heightFt = row.getOrDefault("Height (ft)", "0");
        double heightMeters = 0;
        try { heightMeters = Double.parseDouble(heightFt) * 0.3048; } catch (NumberFormatException ignored) {}
        
        StringBuilder sb = new StringBuilder();
        sb.append("<Placemark>\n<name>").append(siteId).append(" (").append(heightFt).append(" ft)</name>\n");
        sb.append("<styleUrl>#site-icon</styleUrl>\n<ExtendedData>\n<SchemaData schemaUrl=\"#SITES_SCHEME_ID\">\n");
        for (Map.Entry<String, String> entry : row.entrySet()) {
            sb.append("<SimpleData name=\"").append(entry.getKey().replaceAll("[^a-zA-Z0-9]", "")).append("\">").append(entry.getValue()).append("</SimpleData>\n");
        }
        sb.append("</SchemaData>\n</ExtendedData>\n<Point>\n<extrude>1</extrude>\n<altitudeMode>relativeToGround</altitudeMode>\n");
        sb.append("<coordinates>").append(lon).append(",").append(lat).append(",").append(heightMeters).append("</coordinates>\n");
        sb.append("</Point>\n</Placemark>\n");
        return sb.toString();
    }

    private String createSectorPlacemark(Map<String, String> row, String bandName, int range, int bandIndex) {
        StringBuilder sb = new StringBuilder();
        try {
            double lat = Double.parseDouble(row.getOrDefault("Latitude", "0"));
            double lon = Double.parseDouble(row.getOrDefault("Longitude", "0"));
            double azimuth = Double.parseDouble(row.getOrDefault("Azimuth", "0"));
            double height = (Double.parseDouble(row.getOrDefault("Height (ft)", "0")) * 0.3048) + (bandIndex * 0.1); // Add 10cm offset per band
            
            sb.append("<Placemark>\n<name>").append(row.getOrDefault("Custom: NR_Cell_Name", "N/A")).append("</name>\n");
            sb.append("<styleUrl>#").append(bandName.replaceAll("[^a-zA-Z0-9]", "")).append("</styleUrl>\n");
            sb.append("<ExtendedData>\n<SchemaData schemaUrl=\"#SECTORS_SCHEME_ID\">\n");
            for (Map.Entry<String, String> entry : row.entrySet()) {
                sb.append("<SimpleData name=\"").append(entry.getKey().replaceAll("[^a-zA-Z0-9]", "")).append("\">").append(entry.getValue()).append("</SimpleData>\n");
            }
            sb.append("</SchemaData>\n</ExtendedData>\n");
            sb.append("<Polygon>\n<altitudeMode>relativeToGround</altitudeMode>\n<outerBoundaryIs>\n<LinearRing>\n<coordinates>\n");
            sb.append(lon).append(",").append(lat).append(",").append(height).append("\n");
            double beamwidth = 65.0;
            for (int i = 0; i <= 10; i++) {
                double angle = azimuth - (beamwidth / 2) + (beamwidth * i / 10);
                double[] newCoords = getDestinationPoint(lat, lon, angle, range);
                sb.append(newCoords[1]).append(",").append(newCoords[0]).append(",").append(height).append("\n");
            }
            sb.append(lon).append(",").append(lat).append(",").append(height).append("\n");
            sb.append("</coordinates>\n</LinearRing>\n</outerBoundaryIs>\n</Polygon>\n</Placemark>\n");
        } catch (NumberFormatException e) {
            System.err.println("Could not parse number for placemark: " + row.get("Sector ID"));
        }
        return sb.toString();
    }

    private String createLabelPlacemark(Map<String, String> row, String header, int range) {
        StringBuilder sb = new StringBuilder();
        try {
            String labelText = row.getOrDefault(header, "");
            if (!labelText.isEmpty()) {
                double lat = Double.parseDouble(row.getOrDefault("Latitude", "0"));
                double lon = Double.parseDouble(row.getOrDefault("Longitude", "0"));
                double azimuth = Double.parseDouble(row.getOrDefault("Azimuth", "0"));
                double height = Double.parseDouble(row.getOrDefault("Height (ft)", "0")) * 0.3048;
                double distance = header.equals("Electrical Tilt") ? range : range / 2.0;
                double[] labelCoords = getDestinationPoint(lat, lon, azimuth, distance);
                sb.append("<Placemark>\n<name>").append(labelText).append("</name>\n<styleUrl>#label-style</styleUrl>\n");
                sb.append("<ExtendedData>\n<SchemaData schemaUrl=\"#SECTORS_SCHEME_ID\">\n");
                sb.append("<SimpleData name=\"PhysicalCellID\">").append(row.getOrDefault("Physical Cell ID", "")).append("</SimpleData>\n");
                sb.append("<SimpleData name=\"Heightft\">").append(row.getOrDefault("Height (ft)", "")).append("</SimpleData>\n");
                sb.append("<SimpleData name=\"ElectricalTilt\">").append(row.getOrDefault("Electrical Tilt", "")).append("</SimpleData>\n");
                sb.append("</SchemaData>\n</ExtendedData>\n<Point>\n<altitudeMode>relativeToGround</altitudeMode>\n");
                sb.append("<coordinates>").append(labelCoords[1]).append(",").append(labelCoords[0]).append(",").append(height).append("</coordinates>\n");
                sb.append("</Point>\n</Placemark>\n");
            }
        } catch (NumberFormatException e) {
            System.err.println("Could not parse number for label placemark: " + row.get("Sector ID"));
        }
        return sb.toString();
    }

    private double[] getDestinationPoint(double lat, double lon, double bearing, double distance) {
        double R = 6371e3;
        double latRad = Math.toRadians(lat);
        double lonRad = Math.toRadians(lon);
        double bearingRad = Math.toRadians(bearing);
        double lat2Rad = Math.asin(Math.sin(latRad) * Math.cos(distance / R) + Math.cos(latRad) * Math.sin(distance / R) * Math.cos(bearingRad));
        double lon2Rad = lonRad + Math.atan2(Math.sin(bearingRad) * Math.sin(distance / R) * Math.cos(latRad), Math.cos(distance / R) - Math.sin(latRad) * Math.sin(lat2Rad));
        return new double[]{Math.toDegrees(lat2Rad), Math.toDegrees(lon2Rad)};
    }

    public SheetData processSheetWithSAX(File file, String sheetNameToProcess) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(file.getPath())) {
            XSSFReader r = new XSSFReader(pkg);
            SharedStringsTable sst = (SharedStringsTable) r.getSharedStringsTable();
            XMLReader parser = XMLReaderFactory.createXMLReader();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) r.getSheetsData();
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    if (iter.getSheetName().equalsIgnoreCase(sheetNameToProcess)) {
                        SheetContentHandler handler = new SheetContentHandler(sst);
                        parser.setContentHandler(handler);
                        parser.parse(new InputSource(stream));
                        return new SheetData(handler.getHeaders(), handler.getTableData());
                    }
                }
            }
            System.err.println("Sheet '" + sheetNameToProcess + "' not found.");
            return null;
        }
    }

    private DefaultTableModel createTableModel(SheetData sheetData) {
        if (sheetData == null || sheetData.tableData.isEmpty()) return new DefaultTableModel();
        Vector<String> columnHeaders = new Vector<>(sheetData.headers);
        Vector<Vector<Object>> dataVector = new Vector<>();
        for (Map<String, String> rowMap : sheetData.tableData) {
            Vector<Object> rowVector = new Vector<>();
            for (String header : columnHeaders) {
                rowVector.add(rowMap.get(header));
            }
            dataVector.add(rowVector);
        }
        return new DefaultTableModel(dataVector, columnHeaders);
    }

    private static class SheetContentHandler extends DefaultHandler {
        private final SharedStringsTable sst;
        private String lastContents;
        private boolean nextIsString;
        private final List<String> headers = new ArrayList<>();
        private final List<String> currentRow = new ArrayList<>();
        private final List<Map<String, String>> tableData = new ArrayList<>();
        private int currentCellColumn = -1;

        private SheetContentHandler(SharedStringsTable sst) { this.sst = sst; }
        public List<Map<String, String>> getTableData() { return tableData; }
        public List<String> getHeaders() { return headers; }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) {
            if (name.equals("row")) {
                currentCellColumn = -1;
                currentRow.clear();
            } else if (name.equals("c")) {
                currentCellColumn = getColumnIndex(attributes.getValue("r"));
                nextIsString = "s".equals(attributes.getValue("t"));
            }
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name) {
            if (name.equals("v")) {
                if (nextIsString) {
                    try {
                        int idx = Integer.parseInt(lastContents);
                        lastContents = new XSSFRichTextString(sst.getItemAt(idx).getString()).toString();
                    } catch (NumberFormatException e) {
                        System.err.println("SAX Parser Warning: Could not parse shared string index '" + lastContents + "'.");
                    }
                }
                while (currentRow.size() <= currentCellColumn) {
                    currentRow.add("");
                }
                currentRow.set(currentCellColumn, lastContents.trim());
            } else if (name.equals("row")) {
                if (headers.isEmpty() && !currentRow.stream().allMatch(String::isEmpty)) {
                    headers.addAll(currentRow.stream().map(String::trim).collect(Collectors.toList()));
                } else if (!headers.isEmpty()) {
                    Map<String, String> rowMap = new LinkedHashMap<>();
                    for (int i = 0; i < headers.size(); i++) {
                        rowMap.put(headers.get(i), i < currentRow.size() ? currentRow.get(i) : "");
                    }
                    tableData.add(rowMap);
                }
                currentRow.clear();
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            lastContents += new String(ch, start, length);
        }

        private int getColumnIndex(String cellReference) {
            if (cellReference == null) return -1;
            String colRef = cellReference.replaceAll("\\d+", "");
            int colIndex = 0;
            for (int i = 0; i < colRef.length(); i++) {
                colIndex = colIndex * 26 + (colRef.charAt(i) - 'A' + 1);
            }
            return colIndex - 1;
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            PlanetKMLCreator viewer = new PlanetKMLCreator();
            viewer.setVisible(true);
        });
    }
}
