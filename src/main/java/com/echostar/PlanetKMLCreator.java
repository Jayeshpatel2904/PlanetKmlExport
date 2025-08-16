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
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.Vector;
import java.util.stream.Collectors;

/**
 * A Java Swing application that allows a user to select a large Excel file and view
 * specific sheets in a JTable using a memory-efficient SAX parser to prevent OutOfMemoryError.
 * This version opens in full-screen, preserves column order, and performs advanced data merging for the Sectors tab.
 */
public class PlanetKMLCreator extends JFrame {

    private final JTabbedPane tabbedPane;
    private final JLabel statusLabel;
    private DefaultTableModel controllersModel; // Added for the new panel
    private SheetData finalSectorsData; // Store the final processed data
    private SheetData finalSiteData; // Store the final processed site data

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
    }

    public PlanetKMLCreator() {
        super("Excel Sheet Viewer (Memory-Efficient)");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setExtendedState(JFrame.MAXIMIZED_BOTH);
        setSize(1280, 800);
        setLocationRelativeTo(null);

        JPanel mainPanel = new JPanel(new BorderLayout(5, 5));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel topPanel = new JPanel(new BorderLayout(10, 10));
        JButton openButton = new JButton("Open Excel File");
        JButton kmlButton = new JButton("Generate KML");
        statusLabel = new JLabel("No file selected. Please open a large .xlsx file.");
        
        JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        buttonPanel.add(openButton);
        buttonPanel.add(kmlButton);
        
        topPanel.add(buttonPanel, BorderLayout.WEST);
        topPanel.add(statusLabel, BorderLayout.CENTER);

        tabbedPane = new JTabbedPane();
        mainPanel.add(topPanel, BorderLayout.NORTH);
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        
        // Add the support email label at the bottom
        JPanel bottomPanel = new JPanel(new FlowLayout(FlowLayout.CENTER));
        JLabel supportLabel = new JLabel("<html><i><b><font color='red'>For Support Email: jayeshkumar.patel@gmail.com</font></b></i></html>");
        bottomPanel.add(supportLabel);
        mainPanel.add(bottomPanel, BorderLayout.SOUTH);

        add(mainPanel);

        // Add the static controllers panel on startup
        tabbedPane.addTab("Controllers", createControllersPanel());

        openButton.addActionListener(e -> openFile());
        kmlButton.addActionListener(e -> generateKML());
    }

    /**
     * Creates the static panel for managing controllers.
     * This panel is added on startup and is independent of the Excel file.
     * @return A JPanel containing the controllers table and controls.
     */
    private JPanel createControllersPanel() {
        JPanel controllerPanel = new JPanel(new BorderLayout(5, 5));
        String[] columnNames = {"Electrical Controller", "Band"};
        Object[][] data = {
            {"R1", "LB Electrical Tilt"},
            {"R2", "LB Electrical Tilt"},
            {"B", "MB Electrical Tilt"},
            {"Controller_617-894_12", "LB Electrical Tilt"},
            {"Controller_617-894_34", "LB Electrical Tilt"},
            {"Controller 1", "LB Electrical Tilt"},
            {"Controller_1695-2690_56", "MB Electrical Tilt"},
            {"Controller_1695-2690_78", "MB Electrical Tilt"},
            {"Controller 2", "MB Electrical Tilt"},
            {"Controller 3", "MB Electrical Tilt"},
            {"Port 1-2", "LB Electrical Tilt"},
            {"Y1", "MB Electrical Tilt"},
            {"Y2", "MB Electrical Tilt"},
            {"Port 3-4", "MB Electrical Tilt"},
            {"Port 5-6", "MB Electrical Tilt"},
            {"Port 1-4", "LB Electrical Tilt"},
            {"Port 5-8", "MB Electrical Tilt"},
            {"Port 9-10", "MB Electrical Tilt"},
            {"Port 3-4", "LB Electrical Tilt"},
            {"Port 7-8", "MB Electrical Tilt"},
            {"R1 LB Controller", "LB Electrical Tilt"},
            {"R2 LB Controller", "LB Electrical Tilt"},
            {"Y1 HB Controller", "MB Electrical Tilt"},
            {"Y2 HB Controller", "MB Electrical Tilt"}
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
            new Thread(() -> loadExcelData(selectedFile)).start();
        }
    }

    /**
     * Loads data from the selected Excel file, performs transformations on the Sectors data,
     * and displays each sheet in a separate tab.
     * @param file The Excel file to load.
     */
    private void loadExcelData(File file) {
        SwingUtilities.invokeLater(() -> {
            for (int i = tabbedPane.getTabCount() - 1; i >= 0; i--) {
                String title = tabbedPane.getTitleAt(i);
                if (!title.equals("Controllers")) {
                    tabbedPane.remove(i);
                }
            }
        });
        
        String[] sheetsToRead = { "Antennas", "Antenna_Electrical_Parameters", "Sectors", "NR_Sector_Carriers", "Sites" };
        Map<String, SheetData> allSheetsData = new HashMap<>();

        try {
            for (String sheetName : sheetsToRead) {
                SheetData sheetData = processSheetWithSAX(file, sheetName);
                if (sheetData != null) {
                    allSheetsData.put(sheetName, sheetData);
                }
            }

            // Process electrical parameters first to get the Band Info
            SheetData processedElectricalParams = processElectricalParametersData(
                allSheetsData.get("Antenna_Electrical_Parameters")
            );
            
            // Process the Site data
            finalSiteData = processSiteData(allSheetsData.get("Sites"), allSheetsData.get("Antennas"));
            
            // Process the Sectors data to create the final view
            finalSectorsData = processSectorsData(
                allSheetsData.get("Sectors"),
                allSheetsData.get("NR_Sector_Carriers"),
                allSheetsData.get("Antennas"),
                processedElectricalParams // Use the processed data
            );

            // Display the Site tab
            if (finalSiteData != null && !finalSiteData.tableData.isEmpty()) {
                DefaultTableModel tableModel = createTableModel(finalSiteData);
                SwingUtilities.invokeLater(() -> {
                    JTable table = new JTable(tableModel);
                    table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
                    JScrollPane scrollPane = new JScrollPane(table);
                    tabbedPane.insertTab("Sites", null, scrollPane, null, tabbedPane.getTabCount());
                });
            }

            // Display only the final Sectors tab
            if (finalSectorsData != null && !finalSectorsData.tableData.isEmpty()) {
                DefaultTableModel tableModel = createTableModel(finalSectorsData);
                SwingUtilities.invokeLater(() -> {
                    JTable table = new JTable(tableModel);
                    table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
                    JScrollPane scrollPane = new JScrollPane(table);
                    tabbedPane.insertTab("Sectors", null, scrollPane, null, tabbedPane.getTabCount());
                });
            } else {
                System.out.println("No data to display for sheet: Sectors");
            }
            
            SwingUtilities.invokeLater(() -> statusLabel.setText("Successfully loaded and processed: " + file.getName()));
        } catch (Exception e) {
            SwingUtilities.invokeLater(() -> statusLabel.setText("Error processing file: " + e.getMessage()));
            e.printStackTrace();
            JOptionPane.showMessageDialog(this, "Failed to process Excel file: \n" + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private SheetData processSiteData(SheetData siteData, SheetData antennasData) {
        if (siteData == null) {
            System.err.println("Cannot process Site data because the sheet is missing.");
            return null;
        }

        // Create a lookup map to get the first height for each site from the Antennas tab
        Map<String, String> heightLookup = new HashMap<>();
        if (antennasData != null) {
            for (Map<String, String> antennaRow : antennasData.tableData) {
                String siteId = antennaRow.getOrDefault("Site ID", "");
                if (!siteId.isEmpty() && !heightLookup.containsKey(siteId)) {
                    heightLookup.put(siteId, antennaRow.getOrDefault("Height (ft)", ""));
                }
            }
        }

        List<String> finalHeaders = Arrays.asList(
            "Site ID", "Longitude", "Latitude", "Site Name", 
            "Custom: Cluster_ID", "Custom: gNodeB_Id", 
            "Custom: gNodeB_Site_Number", "Custom: TAC", "Height (ft)"
        );
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
        if (sectors == null || nrCarriers == null || antennas == null || electricalParams == null) {
            System.err.println("Cannot process Sectors data because one or more source sheets are missing.");
            return null;
        }

        Map<String, String> pciLookup = new HashMap<>();
        for (Map<String, String> row : nrCarriers.tableData) {
            pciLookup.put(row.getOrDefault("Site ID", "") + "||" + row.getOrDefault("Sector ID", ""), row.get("Physical Cell ID"));
        }

        Map<String, Map<String, String>> antennaLookup = new HashMap<>();
        for (Map<String, String> row : antennas.tableData) {
            antennaLookup.put(row.getOrDefault("Site ID", "") + "||" + row.getOrDefault("Antenna ID", ""), row);
        }
        
        // Create a lookup map for electrical tilt using a virtual KEY.
        Map<String, String> electricalTiltLookup = new HashMap<>();
        for (Map<String, String> paramsRow : electricalParams.tableData) {
            String siteId = paramsRow.getOrDefault("Site ID", "");
            String antennaId = paramsRow.getOrDefault("Antenna ID", "");
            String bandInfo = paramsRow.getOrDefault("Band Info", "");
            String key = siteId + antennaId + bandInfo;
            
            if (!key.isEmpty()) {
                String tiltValue = paramsRow.getOrDefault("Electrical Tilt", "");
                electricalTiltLookup.put(key, tiltValue);
            }
        }
        
        List<String> finalHeaders = Arrays.asList(
            "Site ID", "Band Name", "Custom: NR_Cell_Global_ID", "Custom: NR_Cell_Name",
            "Custom: RU_Model", "Sector ID", "Physical Cell ID", "Antenna ID", "Latitude",
            "Longitude", "Antenna File", "Height (ft)", "Azimuth", "Electrical Tilt"
        );
        List<Map<String, String>> processedData = new ArrayList<>();

        for (Map<String, String> sectorRow : sectors.tableData) {
            Map<String, String> newRow = new LinkedHashMap<>();
            String siteId = sectorRow.getOrDefault("Site ID", "");
            String originalSectorId = sectorRow.getOrDefault("Sector ID", "");
            if (originalSectorId.isEmpty() || siteId.isEmpty()) continue;

            String antennaId = "";
            if (!originalSectorId.isEmpty()) {
                char lastChar = originalSectorId.charAt(originalSectorId.length() - 1);
                if (Character.isDigit(lastChar)) {
                    antennaId = String.valueOf(lastChar);
                }
            }
            
            // Logic to create the virtual KEY for lookup
            String bandName = sectorRow.getOrDefault("Band Name", "").toUpperCase();
            String bandInfo = (bandName.startsWith("N29") || bandName.startsWith("N71")) ? "LB Electrical Tilt" : "MB Electrical Tilt";
            String key = siteId + antennaId + bandInfo;
            
            // Look up the Electrical Tilt using the generated KEY
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
                newRow.put("Latitude", "");
                newRow.put("Longitude", "");
                newRow.put("Antenna File", "");
                newRow.put("Height (ft)", "");
                newRow.put("Azimuth", "");
            }
            newRow.put("Electrical Tilt", electricalTilt);
            processedData.add(newRow);
        }
        return new SheetData(finalHeaders, processedData);
    }

    private SheetData processElectricalParametersData(SheetData electricalParams) {
        if (electricalParams == null) {
            System.err.println("Cannot process Antenna_Electrical_Parameters data because the sheet is missing.");
            return null;
        }

        Map<String, String> controllerBandLookup = new HashMap<>();
        for (int i = 0; i < controllersModel.getRowCount(); i++) {
            String controller = ((String) controllersModel.getValueAt(i, 0));
            String band = ((String) controllersModel.getValueAt(i, 1));
            if (!controller.isEmpty()) {
                controllerBandLookup.put(controller, band);
            }
        }

        List<String> finalHeaders = new ArrayList<>(electricalParams.headers);
        if (!finalHeaders.contains("Band Info")) {
            finalHeaders.add("Band Info");
        }
        
        List<Map<String, String>> processedData = new ArrayList<>();
        for (Map<String, String> electricalRow : electricalParams.tableData) {
            Map<String, String> newRow = new LinkedHashMap<>(electricalRow);
            String controller = electricalRow.getOrDefault("Electrical Controller", "");
            String bandInfo = controllerBandLookup.getOrDefault(controller, "");
            newRow.put("Band Info", bandInfo);
            processedData.add(newRow);
        }
        return new SheetData(finalHeaders, processedData);
    }
    
    private void generateKML() {
        if (finalSectorsData == null || finalSectorsData.tableData.isEmpty() || finalSiteData == null || finalSiteData.tableData.isEmpty()) {
            JOptionPane.showMessageDialog(this, "No data in the Sectors or Site tab to generate KML.", "No Data", JOptionPane.WARNING_MESSAGE);
            return;
        }

        // Get unique band names
        Set<String> uniqueBands = new LinkedHashSet<>();
        for (Map<String, String> row : finalSectorsData.tableData) {
            uniqueBands.add(row.getOrDefault("Band Name", "Unknown"));
        }

        // Show customization dialog
        Map<String, BandSettings> bandSettings = showBandCustomizationDialog(uniqueBands);
        if (bandSettings == null) {
            return; // User cancelled
        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Save KML File");
        
        // --- DYNAMIC FILENAME LOGIC ---
        String firstSiteId = finalSiteData.tableData.get(0).getOrDefault("Site ID", "SITE");
        String sitePart = "SITE";
        if (firstSiteId.length() >= 5) {
            sitePart = firstSiteId.substring(2, 5);
        }
        SimpleDateFormat sdf = new SimpleDateFormat("MMddyyyy");
        String datePart = sdf.format(new Date());
        String fileName = sitePart + "_" + datePart + ".kml";
        fileChooser.setSelectedFile(new File(fileName));
        // --- END DYNAMIC FILENAME LOGIC ---

        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            public boolean accept(File f) {
                return f.getName().toLowerCase().endsWith(".kml") || f.isDirectory();
            }
            public String getDescription() {
                return "KML Files (*.kml)";
            }
        });

        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            File fileToSave = fileChooser.getSelectedFile();
            try (FileWriter writer = new FileWriter(fileToSave)) {
                writer.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n");
                writer.write("<kml xmlns=\"http://www.opengis.net/kml/2.2\">\n");
                writer.write("<Document>\n");

                // Write styles
                writer.write(getSiteStyle()); // Add the detailed site style
                writer.write("<Style id=\"label-style\"><IconStyle><scale>0</scale></IconStyle><LabelStyle><color>ffffffff</color><scale>0.8</scale></LabelStyle></Style>\n");
                for (Map.Entry<String, BandSettings> entry : bandSettings.entrySet()) {
                    if (entry.getValue().include) {
                        writer.write(createKMLStyle(entry.getKey(), entry.getValue().color));
                    }
                }

                // --- SITES FOLDER ---
                writer.write("<Folder>\n<name>SITES</name>\n");
                for (Map<String, String> siteRow : finalSiteData.tableData) {
                    writer.write(createSitePlacemark(siteRow));
                }
                writer.write("</Folder>\n");


                // --- SECTORS FOLDER ---
                writer.write("<Folder>\n<name>SECTORS</name>\n");
                Map<String, List<Map<String, String>>> sectorsByBand = finalSectorsData.tableData.stream()
                    .collect(Collectors.groupingBy(row -> row.getOrDefault("Band Name", "Unknown")));

                for(Map.Entry<String, List<Map<String, String>>> bandEntry : sectorsByBand.entrySet()) {
                    String bandName = bandEntry.getKey();
                    BandSettings settings = bandSettings.get(bandName);
                    if (settings != null && settings.include) {
                        writer.write("<Folder>\n<name>" + bandName + "</name>\n");
                        for (Map<String, String> row : bandEntry.getValue()) {
                            writer.write(createSectorPlacemark(row, bandName, settings.size));
                        }
                        writer.write("</Folder>\n");
                    }
                }
                writer.write("</Folder>\n");
                
                // --- DISPLAY FOLDER ---
                writer.write("<Folder>\n<name>Display</name>\n");
                List<String> displayHeaders = Arrays.asList("Physical Cell ID", "Electrical Tilt");
                for (String header : displayHeaders) {
                    writer.write("<Folder>\n<name>" + header + "</name>\n");
                    
                    Map<String, List<Map<String, String>>> bandsByHeader = finalSectorsData.tableData.stream()
                        .collect(Collectors.groupingBy(row -> row.getOrDefault("Band Name", "Unknown")));

                    for(Map.Entry<String, List<Map<String, String>>> bandEntry : bandsByHeader.entrySet()) {
                        String bandName = bandEntry.getKey();
                        List<Map<String, String>> rowsForBand = bandEntry.getValue();
                        BandSettings settings = bandSettings.get(bandName);
                        
                        boolean createBandFolder = false;
                        if (header.equals("Electrical Tilt")) {
                            createBandFolder = true; // Always create for Electrical Tilt
                        } else if (bandName.toUpperCase().contains("N71")) {
                            createBandFolder = true; // Create for n71 for PCI and Height
                        }

                        if (createBandFolder && settings != null && settings.include) {
                             writer.write("<Folder>\n<name>" + bandName + "</name>\n");
                             for (Map<String, String> row : rowsForBand) {
                                writer.write(createLabelPlacemark(row, header, settings.size));
                            }
                             writer.write("</Folder>\n");
                        }
                    }
                    writer.write("</Folder>\n");
                }
                writer.write("</Folder>\n");


                writer.write("</Document>\n");
                writer.write("</kml>\n");
                JOptionPane.showMessageDialog(this, "KML file generated successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(this, "Error generating KML file: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private Map<String, BandSettings> showBandCustomizationDialog(Set<String> bands) {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(2, 5, 2, 5);
        gbc.gridx = 0;
        gbc.gridy = 0;

        Map<String, JCheckBox> checkBoxes = new HashMap<>();
        Map<String, JButton> colorButtons = new HashMap<>();
        Map<String, JSpinner> sizeSpinners = new HashMap<>();
        Map<String, BandSettings> settingsMap = new HashMap<>();

        // Add headers
        panel.add(new JLabel("Include"), gbc);
        gbc.gridx++;
        panel.add(new JLabel("Band Name"), gbc);
        gbc.gridx++;
        panel.add(new JLabel("Color"), gbc);
        gbc.gridx++;
        panel.add(new JLabel("Size (m)"), gbc);
        
        gbc.gridy++;
        
        Color[] brightColors = {Color.CYAN, Color.MAGENTA, Color.YELLOW, Color.GREEN, Color.ORANGE, Color.PINK};
        int colorIndex = 0;

        for (String band : bands) {
            BandSettings settings = new BandSettings();
            
            // Set default size based on band name
            String upperBand = band.toUpperCase();
            if (upperBand.contains("N71")) {
                settings.size = 500;
            } else if (upperBand.contains("N70")) {
                settings.size = 400;
            } else if (upperBand.contains("N66")) {
                settings.size = 350;
            } else if (upperBand.contains("N29")) {
                settings.size = 300;
            } else {
                settings.size = 500; // Default for others
            }
            
            // Assign a bright color
            settings.color = brightColors[colorIndex % brightColors.length];
            colorIndex++;
            
            settingsMap.put(band, settings);

            gbc.gridx = 0;
            JCheckBox checkBox = new JCheckBox("", true);
            checkBoxes.put(band, checkBox);
            panel.add(checkBox, gbc);

            gbc.gridx++;
            panel.add(new JLabel(band), gbc);

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
            
            gbc.gridy++;
        }

        int result = JOptionPane.showConfirmDialog(this, panel, "Customize KML Bands", JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (result == JOptionPane.OK_OPTION) {
            for (String band : bands) {
                BandSettings settings = settingsMap.get(band);
                settings.include = checkBoxes.get(band).isSelected();
                settings.size = (int) sizeSpinners.get(band).getValue();
            }
            return settingsMap;
        }
        return null; // User cancelled
    }

    private String getSiteStyle() {
        return "<Style id=\"normPointStyle\">\n" +
               "    <IconStyle>\n" +
               "        <scale>0.8</scale>\n" +
               "        <Icon>\n" +
               "            <href>https://i.ibb.co/5YtdGtG/LOGO-PLOT-TRNS.png</href>\n" +
               "        </Icon>\n" +
               "        <hotSpot x=\"0.5\" y=\"0.5\" xunits=\"fraction\" yunits=\"fraction\"/>\n" +
               "    </IconStyle>\n" +
               "    <LabelStyle>\n" +
               "        <color>ff00ffff</color>\n" +
               "        <scale>0.7</scale>\n" +
               "    </LabelStyle>\n" +
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
               "<Style id=\"site-line\"><LineStyle><color>ffffffff</color><width>15</width></LineStyle></Style>\n";
    }

    private String createKMLStyle(String id, Color color) {
        String safeId = id.replaceAll("[^a-zA-Z0-9]", "");
        StringBuilder sb = new StringBuilder();
        sb.append(String.format("<Style id=\"%s\">\n", safeId));

        String colorHex = String.format("80%02x%02x%02x", color.getBlue(), color.getGreen(), color.getRed()); // 50% transparency
        sb.append("  <LineStyle>\n");
        sb.append("    <color>ff").append(colorHex.substring(2)).append("</color>\n"); // Opaque outline
        sb.append("  </LineStyle>\n");
        sb.append("  <PolyStyle>\n");
        sb.append("    <color>").append(colorHex).append("</color>\n");
        sb.append("  </PolyStyle>\n");

        sb.append("</Style>\n");
        return sb.toString();
    }

    private String createSitePlacemark(Map<String, String> row) {
        StringBuilder sb = new StringBuilder();
        String siteId = row.getOrDefault("Site ID", "N/A");
        String lon = row.getOrDefault("Longitude", "0");
        String lat = row.getOrDefault("Latitude", "0");
        String heightFt = row.getOrDefault("Height (ft)", "0");
        double heightMeters = 0;
        try {
            heightMeters = Double.parseDouble(heightFt) * 0.3048; // Convert feet to meters
        } catch (NumberFormatException e) {
            // Keep height as 0 if parsing fails
        }

        // Placemark for the line (mast)
        sb.append("<Placemark>\n");
        sb.append("  <name>").append(siteId).append(" Mast</name>\n");
        sb.append("  <styleUrl>#site-line</styleUrl>\n");
        sb.append("  <LineString>\n");
        sb.append("    <extrude>1</extrude>\n");
        sb.append("    <altitudeMode>relativeToGround</altitudeMode>\n");
        sb.append("    <coordinates>");
        sb.append(lon).append(",").append(lat).append(",0 "); // Start at ground
        sb.append(lon).append(",").append(lat).append(",").append(heightMeters); // End at height
        sb.append("</coordinates>\n");
        sb.append("  </LineString>\n");
        sb.append("</Placemark>\n");

        // Placemark for the icon at the top
        sb.append("<Placemark>\n");
        sb.append("  <name>").append(siteId).append(" (").append(heightFt).append(" ft)</name>\n");
        sb.append("  <styleUrl>#site-icon</styleUrl>\n");
        sb.append("  <description><![CDATA[<table border='1' style='width:100%'>\n");
        for (Map.Entry<String, String> entry : row.entrySet()) {
            sb.append("<tr><td><b>").append(entry.getKey()).append("</b></td><td>").append(entry.getValue()).append("</td></tr>");
        }
        sb.append("  </table>]]></description>\n");
        sb.append("  <Point>\n");
        sb.append("    <altitudeMode>relativeToGround</altitudeMode>\n");
        sb.append("    <coordinates>");
        sb.append(lon).append(",").append(lat).append(",").append(heightMeters);
        sb.append("</coordinates>\n");
        sb.append("  </Point>\n");
        sb.append("</Placemark>\n");

        return sb.toString();
    }

    private String createSectorPlacemark(Map<String, String> row, String bandName, int range) {
        StringBuilder sb = new StringBuilder();
        sb.append("<Placemark>\n");
        sb.append("<name>").append(row.getOrDefault("Sector ID", "N/A")).append("</name>\n");
        sb.append("<styleUrl>#").append(bandName.replaceAll("[^a-zA-Z0-9]", "")).append("</styleUrl>\n");

        sb.append("<description><![CDATA[");
        sb.append("<table border='1' style='width:100%'>");
        for (Map.Entry<String, String> entry : row.entrySet()) {
            sb.append("<tr><td><b>").append(entry.getKey()).append("</b></td><td>").append(entry.getValue()).append("</td></tr>");
        }
        sb.append("</table>]]></description>\n");

        try {
            double lat = Double.parseDouble(row.getOrDefault("Latitude", "0"));
            double lon = Double.parseDouble(row.getOrDefault("Longitude", "0"));
            double azimuth = Double.parseDouble(row.getOrDefault("Azimuth", "0"));
            double height = Double.parseDouble(row.getOrDefault("Height (ft)", "0")) * 0.3048; // Convert feet to meters

            // Create fan shape
            sb.append("<Polygon>\n");
            sb.append("<altitudeMode>relativeToGround</altitudeMode>\n");
            sb.append("<outerBoundaryIs>\n<LinearRing>\n<coordinates>\n");

            sb.append(lon).append(",").append(lat).append(",").append(height).append("\n");

            double beamwidth = 65.0; // Horizontal beamwidth in degrees

            for (int i = 0; i <= 10; i++) {
                double angle = azimuth - (beamwidth / 2) + (beamwidth * i / 10);
                double[] newCoords = getDestinationPoint(lat, lon, angle, range);
                sb.append(newCoords[1]).append(",").append(newCoords[0]).append(",").append(height).append("\n");
            }

            sb.append(lon).append(",").append(lat).append(",").append(height).append("\n");

            sb.append("</coordinates>\n</LinearRing>\n</outerBoundaryIs>\n</Polygon>\n");
            sb.append("</Placemark>\n");

        } catch (NumberFormatException e) {
            System.err.println("Could not parse number for placemark: " + row.get("Sector ID"));
        }
        
        return sb.toString();
    }

    private String createLabelPlacemark(Map<String, String> row, String header, int range) {
        StringBuilder sb = new StringBuilder();
        try {
            double lat = Double.parseDouble(row.getOrDefault("Latitude", "0"));
            double lon = Double.parseDouble(row.getOrDefault("Longitude", "0"));
            double azimuth = Double.parseDouble(row.getOrDefault("Azimuth", "0"));
            double height = Double.parseDouble(row.getOrDefault("Height (ft)", "0")) * 0.3048; // Convert feet to meters
            String labelText = row.getOrDefault(header, "");

            if (!labelText.isEmpty()) {
                double distance = header.equals("Electrical Tilt") ? range : range / 2.0;
                double[] labelCoords = getDestinationPoint(lat, lon, azimuth, distance);
                sb.append("<Placemark>\n");
                sb.append("  <name>").append(labelText).append("</name>\n");
                sb.append("  <styleUrl>#label-style</styleUrl>\n");
                sb.append("  <description><![CDATA[<table border='1'>\n");
                sb.append("    <tr><td><b>Physical Cell ID</b></td><td>").append(row.getOrDefault("Physical Cell ID", "")).append("</td></tr>\n");
                sb.append("    <tr><td><b>Height (ft)</b></td><td>").append(row.getOrDefault("Height (ft)", "")).append("</td></tr>\n");
                sb.append("    <tr><td><b>Electrical Tilt</b></td><td>").append(row.getOrDefault("Electrical Tilt", "")).append("</td></tr>\n");
                sb.append("  </table>]]></description>\n");
                sb.append("  <Point>\n");
                sb.append("    <altitudeMode>relativeToGround</altitudeMode>\n");
                sb.append("    <coordinates>");
                sb.append(labelCoords[1]).append(",").append(labelCoords[0]).append(",").append(height);
                sb.append("</coordinates>\n");
                sb.append("  </Point>\n");
                sb.append("</Placemark>\n");
            }
        } catch (NumberFormatException e) {
            System.err.println("Could not parse number for label placemark: " + row.get("Sector ID"));
        }
        return sb.toString();
    }

    private double[] getDestinationPoint(double lat, double lon, double bearing, double distance) {
        double R = 6371e3; // Earth's radius in meters
        double latRad = Math.toRadians(lat);
        double lonRad = Math.toRadians(lon);
        double bearingRad = Math.toRadians(bearing);

        double lat2Rad = Math.asin(Math.sin(latRad) * Math.cos(distance / R) +
                                   Math.cos(latRad) * Math.sin(distance / R) * Math.cos(bearingRad));
        double lon2Rad = lonRad + Math.atan2(Math.sin(bearingRad) * Math.sin(distance / R) * Math.cos(latRad),
                                             Math.cos(distance / R) - Math.sin(latRad) * Math.sin(lat2Rad));

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
                    String sheetName = iter.getSheetName();
                    if (sheetName.equalsIgnoreCase(sheetNameToProcess)) {
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
                String cellType = attributes.getValue("t");
                nextIsString = (cellType != null && cellType.equals("s"));
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
                        System.err.println("SAX Parser Warning: Could not parse shared string index '" + lastContents + "'. Using value as-is.");
                    }
                }
                while (currentRow.size() < currentCellColumn) {
                    currentRow.add("");
                }
                // Trim all cell values as they are read to prevent whitespace issues
                currentRow.add(lastContents.trim());
            } else if (name.equals("row")) {
                if (headers.isEmpty() && !currentRow.isEmpty()) {
                    // Trim headers as well
                    for (String header : currentRow) {
                        headers.add(header.trim());
                    }
                } else if (!headers.isEmpty()) {
                    Map<String, String> rowMap = new LinkedHashMap<>();
                    for (int i = 0; i < headers.size(); i++) {
                        if (i < currentRow.size()) {
                            rowMap.put(headers.get(i), currentRow.get(i));
                        } else {
                            rowMap.put(headers.get(i), "");
                        }
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
