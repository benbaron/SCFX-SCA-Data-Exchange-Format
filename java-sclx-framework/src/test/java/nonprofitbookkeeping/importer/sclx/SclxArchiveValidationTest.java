package nonprofitbookkeeping.importer.sclx;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class SclxArchiveValidationTest {
    private static final ObjectMapper MAPPER = new ObjectMapper().findAndRegisterModules();
    private static final Path REPO_ROOT = Path.of("..").toAbsolutePath().normalize();

    @Test
    void schemaAndRuleJsonFilesAreWellFormed() throws Exception {
        List<Path> jsonFiles = List.of(
            REPO_ROOT.resolve("SCLX-1.2-specification-package/sclx-1.3-ledger-native-dualrefs.schema.json"),
            REPO_ROOT.resolve("SCLX-1.2-specification-package/sclx-1.3-ledger-native.validator-rules.json"),
            REPO_ROOT.resolve("SCLX-1.2-specification-package/sclx-1.3-full.schema.json")
        );

        for (Path jsonFile : jsonFiles) {
            assertTrue(Files.exists(jsonFile), "Missing JSON file: " + jsonFile);
            JsonNode node = MAPPER.readTree(Files.readString(jsonFile));
            assertNotNull(node, "JSON parse failed for " + jsonFile);
            assertTrue(node.isObject(), "Top-level JSON should be object for " + jsonFile);
        }
    }

    @Test
    void vbaModulesContainExpectedExportImportEntryPoints() throws Exception {
        List<Path> basFiles = List.of(
            REPO_ROOT.resolve("SCLX-VBA-Macro-Package/SCLX_Ledger_IO_v13_reviewed_fixed_2_documented.bas"),
            REPO_ROOT.resolve("SCLX-1.2-specification-package/SCLX_Ledger_IO_v13_with_supplemental_dualrefs.bas")
        );

        for (Path basFile : basFiles) {
            assertTrue(Files.exists(basFile), "Missing VBA module: " + basFile);
            String text = Files.readString(basFile);
            assertTrue(text.contains("Sub ExportSCLX"), "Expected ExportSCLX entry point in " + basFile);
            assertTrue(text.contains("Sub ImportSCLX"), "Expected ImportSCLX entry point in " + basFile);
        }
    }

    @Test
    void sclxDocumentCanDeserializeLedgerNativeShape() throws Exception {
        String minimal = """
            {
              "format": "SCLX",
              "version": "1.3",
              "organization": { "organizationId": "org-test", "name": "Test Org" },
              "transactions": [
                {
                  "transactionId": "t-1",
                  "transactionDate": "2026-01-31",
                  "description": "Test",
                  "lines": [
                    { "lineId": "l1", "accountId": "1000", "debit": 10.00 },
                    { "lineId": "l2", "accountId": "2000", "credit": 10.00 }
                  ]
                }
              ]
            }
            """;

        SclxDocument document = MAPPER.readValue(minimal, SclxDocument.class);
        assertEquals("SCLX", document.format());
        assertEquals("1.3", document.version());
        assertNotNull(document.transactions());
        assertEquals(1, document.transactions().size());
    }
}
