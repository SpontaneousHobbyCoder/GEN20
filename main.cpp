#include <iostream>
#include <fstream>
#include <sstream>
#include <vector>
#include <map>
#include <string>
#include <algorithm>
#include <random>
#include <chrono>
#include <ctime>
#include <stdexcept>
#include <windows.h>
#include <cstdint> // Added for uint32_t

// CRC32 function definition
uint32_t crc32(const std::string& data) {
    uint32_t crc = 0xFFFFFFFF;
    for (char c : data) {
        crc ^= static_cast<uint8_t>(c);
        for (int j = 0; j < 8; ++j) {
            uint32_t mask = -(crc & 1);
            crc = (crc >> 1) ^ (0xEDB88320 & mask);
        }
    }
    return ~crc;
}

std::string trim(const std::string& str) {
    size_t first = str.find_first_not_of(" \t\r\n");
    if (first == std::string::npos) return "";
    size_t last = str.find_last_not_of(" \t\r\n");
    return str.substr(first, last - first + 1);
}

std::string escapeXml(const std::string& text) {
    std::string result;
    for (char c : text) {
        switch (c) {
            case '&': result += "&"; break;
            case '<': result += "<"; break;
            case '>': result += ">"; break;
            case '"': result += '"'; break;
            case '\'': result += "'"; break;
            default: result += c;
        }
    }
    return result;
}

struct Config {
    int totalVariants = 1;
    int questionsPerVariant = 0;
    int mandatoryQuestionsPerVariant = 0;
    int marginLeft = 30, marginRight = 30, marginTop = 20, marginBottom = 20;
    double lineSpacing = 1.5;
    int spaceBefore = 0, spaceAfter = 0;
    std::string headerText = "Заведение X";
    std::string headerFont = "Times New Roman";
    int headerSize = 14;
    bool headerBold = true, headerItalic = false, headerUnderline = false;
    std::string headerAlignment = "center";
    std::string signatureText = "С уважением, администрация";
    std::string signatureFont = "Times New Roman";
    int signatureSize = 14;
    bool signatureBold = false, signatureItalic = true, signatureUnderline = false;
    std::string signatureAlignment = "center";
    std::string questionFont = "Times New Roman";
    int questionSize = 14;
    bool questionBold = false, questionItalic = false, questionUnderline = false;
    std::string questionAlignment = "left";
    std::string variantFont = "Times New Roman";
    int variantSize = 14;
    bool variantBold = false, variantItalic = false, variantUnderline = false;
    std::string variantAlignment = "left";
    bool useNumberSign = false;
    bool lineBetweenVariants = true;
    bool cutSeparator = false;
    bool hasVariantText = true;         // New field: toggle variant text
    std::string variantText = "Вариант"; // New field: custom variant text
    std::vector<std::string> mandatoryQuestions;
    std::vector<std::string> randomQuestions;
};
Config parseConfig(const std::string& filename) {
    std::ifstream file(filename);
    if (!file) throw std::runtime_error("Не удалось открыть config.txt");

    Config config;
    std::map<std::string, std::string> settings;
    std::string line, section, multiLineKey;
    bool inMultiLine = false;

    while (std::getline(file, line)) {
        line = trim(line);
        if (line.empty() && !inMultiLine) continue;

        if (inMultiLine) {
            if (line == "\"\"\"") {
                inMultiLine = false;
                // Remove leading newline if present
                if (!settings[multiLineKey].empty() && settings[multiLineKey][0] == '\n') {
                    settings[multiLineKey] = settings[multiLineKey].substr(1);
                }
            } else {
                if (settings[multiLineKey].empty()) {
                    settings[multiLineKey] = line; // First line without leading \n
                } else {
                    settings[multiLineKey] += "\n" + line; // Subsequent lines with \n
                }
            }
            continue;
        }

        if (line[0] == '#') {
            if (line == "#обязательные вопросы") section = "mandatory";
            else if (line == "#Случайные вопросы") section = "random";
            else section = "";
            continue;
        }

        if (section.empty()) {
            size_t pos = line.find('=');
            if (pos != std::string::npos) {
                std::string key = trim(line.substr(0, pos));
                std::string value = trim(line.substr(pos + 1));
                if (value == "\"\"\"") {
                    inMultiLine = true;
                    multiLineKey = key;
                    settings[key] = "";
                } else {
                    settings[key] = value;
                }
            }
        } else if (section == "mandatory") {
            config.mandatoryQuestions.push_back(line);
        } else if (section == "random") {
            config.randomQuestions.push_back(line);
        }
    }

    // Corrected parsing logic
    if (settings.count("ВсегоВариантов")) config.totalVariants = std::stoi(settings["ВсегоВариантов"]);
    if (settings.count("ВсегоВопросовВарианта")) config.questionsPerVariant = std::stoi(settings["ВсегоВопросовВарианта"]);
    if (settings.count("ОбязательныхВопросовВарианта")) config.mandatoryQuestionsPerVariant = std::stoi(settings["ОбязательныхВопросовВарианта"]);
    if (settings.count("ПолеСлева")) config.marginLeft = std::stoi(settings["ПолеСлева"]);
    if (settings.count("ПолеСправа")) config.marginRight = std::stoi(settings["ПолеСправа"]);
    if (settings.count("ПолеСверху")) config.marginTop = std::stoi(settings["ПолеСверху"]);
    if (settings.count("ПолеСнизу")) config.marginBottom = std::stoi(settings["ПолеСнизу"]);
    if (settings.count("МеждустрочныйИнтервал")) config.lineSpacing = std::stod(settings["МеждустрочныйИнтервал"]);
    if (settings.count("ИнтервалПередАбзацем")) config.spaceBefore = std::stoi(settings["ИнтервалПередАбзацем"]) * 56.7;
    if (settings.count("ИнтервалПослеАбзаца")) config.spaceAfter = std::stoi(settings["ИнтервалПослеАбзаца"]) * 56.7;
    if (settings.count("Заголовок")) config.headerText = settings["Заголовок"];
    if (settings.count("ШрифтЗаголовка")) config.headerFont = settings["ШрифтЗаголовка"];
    if (settings.count("РазмерЗаголовка")) config.headerSize = std::stoi(settings["РазмерЗаголовка"]);
    if (settings.count("ЖирныйЗаголовок")) config.headerBold = settings["ЖирныйЗаголовок"] == "1";
    if (settings.count("КурсивЗаголовок")) config.headerItalic = settings["КурсивЗаголовок"] == "1";
    if (settings.count("ПодчеркнутыйЗаголовок")) config.headerUnderline = settings["ПодчеркнутыйЗаголовок"] == "1";

    int alignCount = 0;
    if (settings.count("ВыравниваниеЗаголовка_ПоЦентру") && settings["ВыравниваниеЗаголовка_ПоЦентру"] == "1") { config.headerAlignment = "center"; alignCount++; }
    if (settings.count("ВыравниваниеЗаголовка_ПоЛевомуКраю") && settings["ВыравниваниеЗаголовка_ПоЛевомуКраю"] == "1") { config.headerAlignment = "left"; alignCount++; }
    if (settings.count("ВыравниваниеЗаголовка_ПоПравомуКраю") && settings["ВыравниваниеЗаголовка_ПоПравомуКраю"] == "1") { config.headerAlignment = "right"; alignCount++; }
    if (settings.count("ВыравниваниеЗаголовка_ПоШирине") && settings["ВыравниваниеЗаголовка_ПоШирине"] == "1") { config.headerAlignment = "both"; alignCount++; }
    if (alignCount != 1) config.headerAlignment = "center";

    if (settings.count("Подпись")) config.signatureText = settings["Подпись"];
    if (settings.count("ШрифтПодписи")) config.signatureFont = settings["ШрифтПодписи"];
    if (settings.count("РазмерПодписи")) config.signatureSize = std::stoi(settings["РазмерПодписи"]);
    if (settings.count("ЖирныйПодписи")) config.signatureBold = settings["ЖирныйПодписи"] == "1";
    if (settings.count("КурсивПодписи")) config.signatureItalic = settings["КурсивПодписи"] == "1";
    if (settings.count("ПодчеркнутыйПодписи")) config.signatureUnderline = settings["ПодчеркнутыйПодписи"] == "1";

    alignCount = 0;
    if (settings.count("ВыравниваниеПодписи_ПоЦентру") && settings["ВыравниваниеПодписи_ПоЦентру"] == "1") { config.signatureAlignment = "center"; alignCount++; }
    if (settings.count("ВыравниваниеПодписи_ПоЛевомуКраю") && settings["ВыравниваниеПодписи_ПоЛевомуКраю"] == "1") { config.signatureAlignment = "left"; alignCount++; }
    if (settings.count("ВыравниваниеПодписи_ПоПравомуКраю") && settings["ВыравниваниеПодписи_ПоПравомуКраю"] == "1") { config.signatureAlignment = "right"; alignCount++; }
    if (settings.count("ВыравниваниеПодписи_ПоШирине") && settings["ВыравниваниеПодписи_ПоШирине"] == "1") { config.signatureAlignment = "both"; alignCount++; }
    if (alignCount != 1) config.signatureAlignment = "center";

    if (settings.count("ШрифтВопросов")) config.questionFont = settings["ШрифтВопросов"];
    if (settings.count("РазмерВопросов")) config.questionSize = std::stoi(settings["РазмерВопросов"]);
    if (settings.count("ЖирныйВопросов")) config.questionBold = settings["ЖирныйВопросов"] == "1";
    if (settings.count("КурсивВопросов")) config.questionItalic = settings["КурсивВопросов"] == "1";
    if (settings.count("ПодчеркнутыйВопросов")) config.questionUnderline = settings["ПодчеркнутыйВопросов"] == "1";

    alignCount = 0;
    if (settings.count("ВыравниваниеВопросов_ПоЦентру") && settings["ВыравниваниеВопросов_ПоЦентру"] == "1") { config.questionAlignment = "center"; alignCount++; }
    if (settings.count("ВыравниваниеВопросов_ПоЛевомуКраю") && settings["ВыравниваниеВопросов_ПоЛевомуКраю"] == "1") { config.questionAlignment = "left"; alignCount++; }
    if (settings.count("ВыравниваниеВопросов_ПоПравомуКраю") && settings["ВыравниваниеВопросов_ПоПравомуКраю"] == "1") { config.questionAlignment = "right"; alignCount++; }
    if (settings.count("ВыравниваниеВопросов_ПоШирине") && settings["ВыравниваниеВопросов_ПоШирине"] == "1") { config.questionAlignment = "both"; alignCount++; }
    if (alignCount != 1) config.questionAlignment = "left";

    if (settings.count("ШрифтВарианта")) config.variantFont = settings["ШрифтВарианта"];
    if (settings.count("РазмерВарианта")) config.variantSize = std::stoi(settings["РазмерВарианта"]);
    if (settings.count("ЖирныйВариант")) config.variantBold = settings["ЖирныйВариант"] == "1";
    if (settings.count("КурсивВариант")) config.variantItalic = settings["КурсивВариант"] == "1";
    if (settings.count("ПодчеркнутыйВариант")) config.variantUnderline = settings["ПодчеркнутыйВариант"] == "1";

    alignCount = 0;
    if (settings.count("ВыравниваниеВарианта_ПоЦентру") && settings["ВыравниваниеВарианта_ПоЦентру"] == "1") { config.variantAlignment = "center"; alignCount++; }
    if (settings.count("ВыравниваниеВарианта_ПоЛевомуКраю") && settings["ВыравниваниеВарианта_ПоЛевомуКраю"] == "1") { config.variantAlignment = "left"; alignCount++; }
    if (settings.count("ВыравниваниеВарианта_ПоПравомуКраю") && settings["ВыравниваниеВарианта_ПоПравомуКраю"] == "1") { config.variantAlignment = "right"; alignCount++; }
    if (settings.count("ВыравниваниеВарианта_ПоШирине") && settings["ВыравниваниеВарианта_ПоШирине"] == "1") { config.variantAlignment = "both"; alignCount++; }
    if (alignCount != 1) config.variantAlignment = "left";

    if (settings.count("ИспользоватьЗнак№")) config.useNumberSign = settings["ИспользоватьЗнак№"] == "1";
    if (settings.count("ЛинияМеждуВариантами")) config.lineBetweenVariants = settings["ЛинияМеждуВариантами"] == "1";
    if (settings.count("РазделительДляРезки")) config.cutSeparator = settings["РазделительДляРезки"] == "1";
    if (settings.count("ЕстьТекстВарианта")) config.hasVariantText = settings["ЕстьТекстВарианта"] == "1";
    if (settings.count("ТекстВарианта")) {
        std::string rawText = settings["ТекстВарианта"];
        if (rawText.size() >= 6 && rawText.substr(0, 3) == "\"\"\"" && rawText.substr(rawText.size() - 3) == "\"\"\"") {
            config.variantText = rawText.substr(3, rawText.size() - 6); // Strip """ from start and end
        } else {
            config.variantText = rawText; // Fallback if no """ found
        }
    }
    return config;
}
std::vector<std::string> selectQuestions(const std::vector<std::string>& pool, int count, std::mt19937& rng) {
    std::vector<std::string> result;
    if (pool.empty()) return result;

    if (pool.size() >= count) {
        std::vector<std::string> temp = pool;
        std::shuffle(temp.begin(), temp.end(), rng);
        result.assign(temp.begin(), temp.begin() + count);
    } else {
        result = pool;
        std::uniform_int_distribution<int> dist(0, pool.size() - 1);
        for (int i = pool.size(); i < count; ++i) {
            result.push_back(pool[dist(rng)]);
        }
    }
    return result;
}

int estimateVariantHeight(const Config& config, const std::vector<std::string>& questions) {
    int headerLines = std::count(config.headerText.begin(), config.headerText.end(), '\n') + 1;
    int signatureLines = std::count(config.signatureText.begin(), config.signatureText.end(), '\n') + 1;
    int totalLines = headerLines + 1 + questions.size() + signatureLines;
    int lineHeight = static_cast<int>(config.lineSpacing * 240 + config.spaceBefore + config.spaceAfter);
    return totalLines * lineHeight;
}

std::string generateDocumentXml(const Config& config) {
    std::ostringstream xml;
    std::random_device rd;
    std::mt19937 rng(rd());
    int lineSpacingTwips = static_cast<int>(240 * config.lineSpacing + 0.5);
    if (lineSpacingTwips <= 0) lineSpacingTwips = 240;

    const int pageHeight = static_cast<int>((297 - config.marginTop - config.marginBottom) * 56.7);
    int currentHeight = 0;

    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    xml << "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\n";
    xml << "    <w:body>\n";

    int randomQuestionsPerVariant = config.questionsPerVariant - config.mandatoryQuestionsPerVariant;

    for (int var = 1; var <= config.totalVariants; ++var) {
        auto mandatory = selectQuestions(config.mandatoryQuestions, config.mandatoryQuestionsPerVariant, rng);
        auto random = selectQuestions(config.randomQuestions, randomQuestionsPerVariant, rng);
        std::vector<std::string> questions = mandatory;
        questions.insert(questions.end(), random.begin(), random.end());

        int variantHeight = estimateVariantHeight(config, questions);

        bool isNewPage = false;
        if (currentHeight + variantHeight > pageHeight && var > 1) {
            xml << "        <w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>\n";
            currentHeight = 0;
            isNewPage = true;
        } else if (var > 1 && !(var == 3 && isNewPage)) {
            if (config.cutSeparator) {
                xml << "        <w:p><w:pPr><w:jc w:val=\"center\"/><w:spacing w:before=\"" << config.spaceBefore
                    << "\" w:after=\"" << config.spaceAfter << "\" w:line=\"" << lineSpacingTwips
                    << "\" w:lineRule=\"auto\"/></w:pPr><w:r><w:t>__                                                            __</w:t></w:r></w:p>\n";
                currentHeight += lineSpacingTwips + config.spaceBefore + config.spaceAfter;
            } else if (config.lineBetweenVariants) {
                xml << "        <w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"6\" w:space=\"1\" w:color=\"auto\"/></w:pBdr></w:pPr></w:p>\n";
                currentHeight += lineSpacingTwips + config.spaceBefore + config.spaceAfter;
            }
        }

        std::istringstream headerStream(config.headerText);
        std::string headerLine;
        while (std::getline(headerStream, headerLine)) {
            xml << "        <w:p><w:pPr><w:jc w:val=\"" << config.headerAlignment << "\"/><w:spacing w:before=\""
                << config.spaceBefore << "\" w:after=\"" << config.spaceAfter << "\" w:line=\"" << lineSpacingTwips
                << "\" w:lineRule=\"auto\"/><w:keepNext w:val=\"1\"/><w:keepLines w:val=\"1\"/></w:pPr><w:r><w:rPr>";
            xml << "<w:rFonts w:ascii=\"" << config.headerFont << "\" w:hAnsi=\"" << config.headerFont << "\"/>";
            xml << "<w:sz w:val=\"" << (config.headerSize * 2) << "\"/>";
            if (config.headerBold) xml << "<w:b w:val=\"1\"/>";
            if (config.headerItalic) xml << "<w:i w:val=\"1\"/>";
            if (config.headerUnderline) xml << "<w:u w:val=\"single\"/>";
            xml << "</w:rPr><w:t>" << escapeXml(headerLine) << "</w:t></w:r></w:p>\n";
        }

        // Generate variant text based on new logic
        if (config.hasVariantText) {
            std::string baseVariantText;
            if (config.useNumberSign) {
                if (config.variantText.empty()) {
                    baseVariantText = "№" + std::to_string(var);
                } else {
                    baseVariantText = config.variantText + " №" + std::to_string(var);
                }
            } else {
                if (config.variantText.empty()) {
                    baseVariantText = std::to_string(var);
                } else {
                    baseVariantText = config.variantText + " " + std::to_string(var);
                }
            }
        
            std::istringstream variantStream(baseVariantText);
            std::string variantLine;
            while (std::getline(variantStream, variantLine)) {
                xml << "        <w:p><w:pPr><w:jc w:val=\"" << config.variantAlignment << "\"/><w:spacing w:before=\""
                    << config.spaceBefore << "\" w:after=\"" << config.spaceAfter << "\" w:line=\"" << lineSpacingTwips
                    << "\" w:lineRule=\"auto\"/><w:keepNext w:val=\"1\"/><w:keepLines w:val=\"1\"/></w:pPr><w:r><w:rPr>";
                xml << "<w:rFonts w:ascii=\"" << config.variantFont << "\" w:hAnsi=\"" << config.variantFont << "\"/>";
                xml << "<w:sz w:val=\"" << (config.variantSize * 2) << "\"/>";
                if (config.variantBold) xml << "<w:b w:val=\"1\"/>";
                if (config.variantItalic) xml << "<w:i w:val=\"1\"/>";
                if (config.variantUnderline) xml << "<w:u w:val=\"single\"/>";
                xml << "</w:rPr><w:t>" << escapeXml(variantLine) << "</w:t></w:r></w:p>\n";
            }
        }

        for (size_t i = 0; i < questions.size(); ++i) {
            xml << "        <w:p><w:pPr><w:jc w:val=\"" << config.questionAlignment << "\"/><w:spacing w:before=\""
                << config.spaceBefore << "\" w:after=\"" << config.spaceAfter << "\" w:line=\"" << lineSpacingTwips
                << "\" w:lineRule=\"auto\"/><w:keepNext w:val=\"1\"/><w:keepLines w:val=\"1\"/></w:pPr><w:r><w:rPr>";
            xml << "<w:rFonts w:ascii=\"" << config.questionFont << "\" w:hAnsi=\"" << config.questionFont << "\"/>";
            xml << "<w:sz w:val=\"" << (config.questionSize * 2) << "\"/>";
            if (config.questionBold) xml << "<w:b w:val=\"1\"/>";
            if (config.questionItalic) xml << "<w:i w:val=\"1\"/>";
            if (config.questionUnderline) xml << "<w:u w:val=\"single\"/>";
            xml << "</w:rPr><w:t>" << (i + 1) << ". " << escapeXml(questions[i]) << "</w:t></w:r></w:p>\n";
        }

        std::istringstream signatureStream(config.signatureText);
        std::string signatureLine;
        while (std::getline(signatureStream, signatureLine)) {
            xml << "        <w:p><w:pPr><w:jc w:val=\"" << config.signatureAlignment << "\"/><w:spacing w:before=\""
                << config.spaceBefore << "\" w:after=\"" << config.spaceAfter << "\" w:line=\"" << lineSpacingTwips
                << "\" w:lineRule=\"auto\"/><w:keepNext w:val=\"1\"/><w:keepLines w:val=\"1\"/></w:pPr><w:r><w:rPr>";
            xml << "<w:rFonts w:ascii=\"" << config.signatureFont << "\" w:hAnsi=\"" << config.signatureFont << "\"/>";
            xml << "<w:sz w:val=\"" << (config.signatureSize * 2) << "\"/>";
            if (config.signatureBold) xml << "<w:b w:val=\"1\"/>";
            if (config.signatureItalic) xml << "<w:i w:val=\"1\"/>";
            if (config.signatureUnderline) xml << "<w:u w:val=\"single\"/>";
            xml << "</w:rPr><w:t>" << escapeXml(signatureLine) << "</w:t></w:r></w:p>\n";
        }

        currentHeight += variantHeight;
    }

    xml << "        <w:sectPr><w:pgMar w:top=\"" << static_cast<int>(config.marginTop * 56.7) << "\" w:right=\""
        << static_cast<int>(config.marginRight * 56.7) << "\" w:bottom=\"" << static_cast<int>(config.marginBottom * 56.7)
        << "\" w:left=\"" << static_cast<int>(config.marginLeft * 56.7) << "\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/><w:cols w:space=\"720\"/></w:sectPr>\n";
    xml << "    </w:body>\n";
    xml << "</w:document>\n";

    return xml.str();
}
void createDocx(const std::string& filename, const std::string& documentXml) {
    std::ofstream ofs(filename, std::ios::binary);
    if (!ofs) throw std::runtime_error("Не удалось создать DOCX-файл");

    std::string contentTypesXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                                  "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
                                  "    <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
                                  "    <Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>\n"
                                  "</Types>";
    std::string relsXml = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
                          "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
                          "    <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>\n"
                          "</Relationships>";

    struct FileEntry {
        std::string name;
        std::string data;
        uint32_t offset;
    };
    std::vector<FileEntry> files = {
        {"[Content_Types].xml", contentTypesXml, 0},
        {"_rels/.rels", relsXml, 0},
        {"word/document.xml", documentXml, 0}
    };

    uint32_t centralDirOffset = 0;
    for (auto& file : files) {
        file.offset = static_cast<uint32_t>(ofs.tellp());
        uint32_t crc = crc32(file.data); // Now compiles because crc32 is defined above
        uint32_t size = file.data.size();
        uint16_t nameLen = file.name.size();

        ofs.write("\x50\x4B\x03\x04", 4);
        ofs.write("\x14\x00", 2);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00\x00\x00", 4);
        ofs.write(reinterpret_cast<const char*>(&crc), 4);
        ofs.write(reinterpret_cast<const char*>(&size), 4);
        ofs.write(reinterpret_cast<const char*>(&size), 4);
        ofs.write(reinterpret_cast<const char*>(&nameLen), 2);
        ofs.write("\x00\x00", 2);
        ofs.write(file.name.c_str(), nameLen);
        ofs.write(file.data.c_str(), size);
    }

    centralDirOffset = static_cast<uint32_t>(ofs.tellp());
    for (const auto& file : files) {
        uint32_t crc = crc32(file.data); // Now compiles because crc32 is defined above
        uint32_t size = file.data.size();
        uint16_t nameLen = file.name.size();

        ofs.write("\x50\x4B\x01\x02", 4);
        ofs.write("\x14\x00\x14\x00", 4);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00\x00\x00", 4);
        ofs.write(reinterpret_cast<const char*>(&crc), 4);
        ofs.write(reinterpret_cast<const char*>(&size), 4);
        ofs.write(reinterpret_cast<const char*>(&size), 4);
        ofs.write(reinterpret_cast<const char*>(&nameLen), 2);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00", 2);
        ofs.write("\x00\x00\x00\x00", 4);
        ofs.write(reinterpret_cast<const char*>(&file.offset), 4);
        ofs.write(file.name.c_str(), nameLen);
    }

    uint32_t centralDirSize = static_cast<uint32_t>(ofs.tellp()) - centralDirOffset;
    uint16_t fileCount = files.size();
    ofs.write("\x50\x4B\x05\x06", 4);
    ofs.write("\x00\x00", 2);
    ofs.write("\x00\x00", 2);
    ofs.write(reinterpret_cast<const char*>(&fileCount), 2);
    ofs.write(reinterpret_cast<const char*>(&fileCount), 2);
    ofs.write(reinterpret_cast<const char*>(&centralDirSize), 4);  // Fixed typo
    ofs.write(reinterpret_cast<const char*>(&centralDirOffset), 4); // Fixed typo
    ofs.write("\x00\x00", 2);
}

int main() {
    SetConsoleOutputCP(CP_UTF8);
    try {
        Config config = parseConfig("config.txt");

        if (config.questionsPerVariant < config.mandatoryQuestionsPerVariant) {
            std::wstring errorMsg = L"Ошибка: ВсегоВопросовВарианта < ОбязательныхВопросовВарианта";
            MessageBoxW(NULL, errorMsg.c_str(), L"Ошибка конфигурации", MB_OK | MB_ICONERROR);
            return 1;
        }
        if (config.mandatoryQuestions.size() < config.mandatoryQuestionsPerVariant) {
            std::wstring errorMsg = L"Ошибка: Недостаточно обязательных вопросов. Требуется " +
                                    std::to_wstring(config.mandatoryQuestionsPerVariant) +
                                    L", доступно " + std::to_wstring(config.mandatoryQuestions.size()) + L".";
            MessageBoxW(NULL, errorMsg.c_str(), L"Ошибка конфигурации", MB_OK | MB_ICONERROR);
            return 1;
        }

        int randomQuestionsPerVariant = config.questionsPerVariant - config.mandatoryQuestionsPerVariant;
        if (config.randomQuestions.size() < randomQuestionsPerVariant) {
            int shortage = randomQuestionsPerVariant - config.randomQuestions.size();
            std::wstring warningMsg = L"В config.txt недостаточно случайных вопросов.\nТребуется: " +
                                      std::to_wstring(randomQuestionsPerVariant) +
                                      L", доступно: " + std::to_wstring(config.randomQuestions.size()) +
                                      L".\nНекоторые вопросы будут повторяться.\nРекомендуется добавить еще " +
                                      std::to_wstring(shortage) + L" случайных вопросов.";
            MessageBoxW(NULL, warningMsg.c_str(), L"Предупреждение", MB_OK | MB_ICONWARNING);
        }

        auto now = std::chrono::system_clock::now();
        std::time_t now_time = std::chrono::system_clock::to_time_t(now);
        std::tm* now_tm = std::localtime(&now_time);
        char buffer[20];
        std::strftime(buffer, 20, "%H.%M.%S_%d.%m.%Y", now_tm);
        std::string filename = std::string(buffer) + ".docx";

        std::string documentXml = generateDocumentXml(config);
        createDocx(filename, documentXml);

        std::cout << "DOCX-файл успешно создан: " << filename << "\n";
        HINSTANCE result = ShellExecute(NULL, "open", filename.c_str(), NULL, NULL, SW_SHOWNORMAL);
        if ((intptr_t)result <= 32) {
            std::cerr << "Ошибка: Не удалось открыть " << filename << ". Код ошибки: " << (intptr_t)result << "\n";
        }
    } catch (const std::exception& e) {
        std::wstring errorMsg = L"Ошибка: " + std::wstring(e.what(), e.what() + strlen(e.what()));
        MessageBoxW(NULL, errorMsg.c_str(), L"Критическая ошибка", MB_OK | MB_ICONERROR);
        return 1;
    }
    return 0;
}