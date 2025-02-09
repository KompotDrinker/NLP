import docx2txt
import PyPDF2 
import re
import spacy
import pandas as pd
from spacy.matcher import Matcher
from spacy.language import Language
import os
import constants
import json


class Parser:

    def __init__(self, nlp):
        self.nlp = nlp
        self.ruler = self.nlp.add_pipe("entity_ruler", before="ner")

    def doctotext(self, m):
        temp = docx2txt.process(m)
        resume_text = []
        for line in temp.split('\n'):
            if line:  # Ignore empty lines
                resume_text.append(line.replace('\t', ' '))
        text = "\n".join(resume_text)
        return (text)

    def pdftotext(self,m):
        pdfFileObj = open(m, 'rb')
        pdfFileReader = PyPDF2.PdfFileReader(pdfFileObj)
        num_pages = pdfFileReader.numPages
        currentPageNumber = 0
        text = ''
        while(currentPageNumber < num_pages ):
            pdfPage = pdfFileReader.getPage(currentPageNumber)
            text = text + pdfPage.extractText()
            currentPageNumber += 1
        return (text)

    def extract_mobile_number(self,resume_text):
        phone = re.findall(re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]'), resume_text)
        if phone:
            number = ''.join(phone[0])
            return number

    def extract_email_addresses(self,string):
        r = re.compile(r'[\w\.-]+@[\w\.-]+')
        return r.findall(string)

    def extract_name(self, resume_text):
        for line in resume_text.split('\n'):
            print(line+"\n")
            nlp_text = self.nlp(line)
            for ent in nlp_text.ents:
                if(ent.label_=="PER"):
                    return ent

    def extract_skills(self,resume_text):
        data = pd.read_csv('skills.csv') 
        skills = data[data.columns[0]].tolist()
        skills=list(map(str,skills))

        skill_patterns = [{"label": "SKILL_TITLE", "pattern": title} for title in skills]
        self.ruler.add_patterns(skill_patterns)

        nlp_text = self.nlp(resume_text)
        
        skillset = []

        for ent in nlp_text.ents:
                print
                if(ent.label_=="SKILL_TITLE"):
                    skillset.append(ent.text)
        print(skillset)


        return list(set(skillset))

    def contains_word(self,word, word_list):
        return word.lower() in (w.lower() for w in word_list)


    def extract_experience(self, resume_text):
        positions = []
        companies = []
        experience = []
        
        

        if "ner" in self.nlp.pipe_names:
            self.nlp.remove_pipe("ner")
        job_patterns = [{"label": "JOB_TITLE", "pattern": title.lower()} for title in constants.TITLES]
        self.ruler.add_patterns(job_patterns)
        company_patterns = [{"label": "COMPANY_TITLE", "pattern": title.lower()} for title in constants.COMPANIES]
        self.ruler.add_patterns(company_patterns)
        self.nlp.add_pipe("custom_date_finder", last=True)

        text=resume_text.lower()
        lines=text.split('\n')

        for indx, line in enumerate(lines):
            nlp_text = self.nlp(line)
            for ent in nlp_text.ents:
                if ent.label_ == "JOB_TITLE":  # Посада
                    val=self.find_entity_within_radius(lines,indx,"COMPANY_TITLE")
                    if val!=[]:
                        positions.append(ent.text)
                        companies.append(val[0])
                        date=self.find_entity_within_radius(lines,indx,"YEAR") 
                        if len(date)>1:
                            experience.append(date[0]+" - "+date[1])
                        elif len(date)<1:
                            experience.append(None)
                        else:
                            experience.append(date[0])
                        

        for i in range(min(len(positions), len(companies), len(experience))):
            experience.append({
                "position": positions[i],
                "company": companies[i],
                "years": experience[i]
            })

        return experience

    def find_entity_within_radius(self,lines, target_index, pattern, radius=2):
        Values=[]
        for i in range(max(0, target_index - radius), min(len(lines), target_index + radius + 1)):
            nlp_text = self.nlp(lines[i])

            for ent in nlp_text.ents:
                if ent.label_ == pattern:  # Посада
                    Values.append(ent.text)
        
        return Values

    def get_education(self,document):
        output = []
        education = []
        institution = []
        for line in document.lower().split('\n'):
            for word in line.split(' '):
                if len(word) > 2 and word in constants.EDUCATION:
                    if line not in education:
                        education.append(line)
                if len(word) > 2 and (word in constants.INSTITUTION or self.contains_word(word,constants.INSTITUTION)):
                    if line not in institution:
                        institution.append(line)
        for i in range(min(len(education), len(institution))):
            output.append({
                "Education": education[i],
                "Institution": institution[i],
            })
        return education, institution

    @Language.component("custom_date_finder")
    def custom_date_finder(doc):
        matches = re.compile(r"(?<!\d)(19\d{2}|20\d{2})(?!\d)").finditer(doc.text)
        new_ents = []
        for match in matches:
            start, end = match.span()
            span = doc.char_span(start, end, label="YEAR")
            if span is not None:
                new_ents.append(span)
        doc.ents = list(doc.ents) + new_ents  # Додаємо знайдені дати у doc.ents
        return doc
    def parse_resume(self,link):
        FilePath = link
        FilePath.endswith(('.pdf', '.docx'))
        # textinput
        if FilePath.endswith('.docx'):
            textinput = self.doctotext(FilePath) 
        elif FilePath.endswith('.pdf'):
            textinput = self.pdftotext(FilePath)
        else:
            print("File not supported")
            return None
        name_match = self.extract_name(textinput)
        phone_match = self.extract_mobile_number(textinput)
        email_match = self.extract_email_addresses(textinput)

        return {
        "name": name_match if name_match else None,
        "contact_info": {
            "phone": phone_match if phone_match else None,
            "email": email_match if email_match else None,
        },
        "skills": self.extract_skills(textinput),
        "experience": self.extract_experience(textinput),
        "education": self.get_education(textinput)
        
    }

    def parse_folder(self,link):
        parsed_resumes = []
        for file in os.listdir(link):
            file_path = os.path.join(link, file)
            parsed_resumes.append(self.parse_resume(file_path))
        with open("parsed_resumes.json", "w", encoding="utf-8") as f:
            json.dump({"resumes": parsed_resumes}, f, ensure_ascii=False, indent=4)


nlp = spacy.load('uk_core_news_sm')
parser=Parser(nlp)
print(parser.parse_resume("CVs\Big_Data_Software_Engineer_CV.docx"))