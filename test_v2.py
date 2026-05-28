from populate import populate_template

data = {
    "name": "Jane Smith",
    "position": "Candidate Profile",
    "summary": "Jane is a senior product manager with twelve years of experience across fintech and enterprise SaaS. She has a strong track record of launching zero-to-one products and scaling them to eight-figure revenue. Known for her ability to align cross-functional teams and translate complex technical requirements into clear product strategy.",
    "achievement": {
        "title": "Scaled payments platform to 4M daily transactions",
        "bullet1": "Delivered 99.99% uptime across 14 markets in 18 months",
        "bullet2": "Reduced onboarding time by 60% through UX redesign programme",
        "bullet3": "Led team of 12 PMs and 40 engineers across three time zones"
    },
    "experience": [
        {
            "startYear": "2020", "endYear": "Present",
            "company": "Stripe", "position": "Senior Product Manager",
            "bullet1": "Owned the global payments routing product serving 4M daily transactions across 14 markets.",
            "bullet2": "Led cross-functional squad of 40 engineers to deliver 99.99% uptime SLA.",
            "bullet3": "Reduced merchant onboarding time by 60% through end-to-end UX redesign."
        },
        {
            "startYear": "2017", "endYear": "2020",
            "company": "Monzo", "position": "Product Manager",
            "bullet1": "Built the savings and investment feature from zero to 800k active users.",
            "bullet2": "Drove 35% increase in daily active usage through personalisation engine.",
            "bullet3": None
        }
    ],
    "education": [
        {
            "startYear": "2010", "endYear": "2013",
            "institution": "London School of Economics",
            "degree": "BSc Economics",
            "bullet1": "First Class Honours",
            "bullet2": None
        }
    ],
    "languages": [
        {"name": "English", "fluency": "Native"},
        {"name": "French",  "fluency": "Fluent"}
    ],
    "skills": ["Product Strategy", "Roadmap Planning", "A/B Testing", "SQL", "Figma", "Agile"],
    "certifications": [
        {"name": "AWS Certified Solutions Architect", "issuer": "Amazon Web Services", "year": "2022"},
        {"name": "Professional Scrum Master I", "issuer": "Scrum.org", "year": "2021"},
        {"name": "Google Analytics Individual Qualification", "issuer": "Google", "year": None}
    ]
}

populate_template(
    r"C:\Users\W.Dier\Downloads\pm-cv-formatter-app\assets\PM_CV_Template_v2.docx",
    data,
    r"C:\Users\W.Dier\Downloads\pm-cv-formatter-app\outputs\TEST_v2.docx"
)
print("SUCCESS - outputs/TEST_v2.docx")
