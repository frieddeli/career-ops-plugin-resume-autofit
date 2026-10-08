"""
example_facts.py
----------------
Sample candidate profile data demonstrating how to structure facts for the ResumeBuilder engine.
Replace this with your own facts module or database queries.
"""

PROFILE_EN = {
    "name": "Alex Taylor",
    "mobile": "+1 (555) 019-2834",
    "email": "alex.taylor@example.com",
    "links": [
        ("LinkedIn", "https://linkedin.com/in/example-alex"),
        ("GitHub", "https://github.com/example-alex"),
        ("Portfolio", "https://example.com"),
    ],
    "badge": "Open Source Contributor",
    "sub_badge": "Cloud Architecture & Infrastructure Engineering",
    "summary_title": "Professional Summary",
    "summary": "Cloud Solutions and Systems Engineer with experience designing scalable distributed architectures, optimizing edge routing, and building enterprise data telemetry pipelines. Proficient in Python, Go, Kubernetes, and Terraform. Passionate about reliability engineering and developer tooling.",
    "education": {
        "title": "Education",
        "school": "State University of Technology",
        "period": "Aug 2021 – May 2025",
        "degree": "B.S. in Computer Science & Engineering (Honors) | GPA: 3.85 / 4.0",
        "details": [
            ("Relevant Coursework", "Distributed Systems, Cloud Computing, Operating Systems, Computer Networks, Database Internals."),
            ("Capstone Project", "High-throughput asynchronous event ingestion platform with sub-millisecond p99 latency."),
        ]
    },
    "experience": {
        "title": "Work Experience",
        "roles": [
            {
                "company": "CloudScale Technologies",
                "period": "Jun 2024 – Present",
                "role": "Site Reliability Engineering Intern, Infrastructure Operations",
                "bullets": [
                    ("Distributed Storage Observability", "Architected an end-to-end telemetry pipeline using Prometheus and Grafana to track query latencies across distributed storage nodes, reducing diagnostic MTTR by 35%."),
                    ("Edge Network Optimization", "Diagnosed cross-region routing degradation across reverse proxies and DNS resolvers, tuning TCP keep-alives and buffer allocations for a 20% throughput uplift."),
                    ("Automated Incident Runbooks", "Authored Ansible and Python automation playbooks for canary rollouts, eliminating manual configuration drift across 200+ staging nodes."),
                ]
            },
            {
                "company": "Datacenter Dynamics Inc.",
                "period": "May 2023 – Aug 2023",
                "role": "Systems Engineering Intern, Fleet Reliability",
                "bullets": [
                    ("Hardware Telemetry Pipeline", "Aggregated multi-sensor hardware logs using BigQuery SQL and Looker dashboards to identify transceivers nearing mean-time-to-failure (MTTF)."),
                    ("Incident Triage Automation", "Built a Python-based incident triage CLI integrating with PagerDuty and Slack to automatically classify alert severity and page appropriate on-call squads."),
                ]
            }
        ]
    },
    "skills": {
        "title": "Technical Skills",
        "categories": [
            ("Cloud & Infrastructure", "Kubernetes, Docker, Terraform, AWS, Google Cloud, Linux, Ansible, Envoy"),
            ("Languages & Databases", "Python, Go, Bash, SQL, PostgreSQL, Redis, ClickHouse"),
            ("Reliability & Tools", "Prometheus, Grafana, OpenTelemetry, CI/CD (GitHub Actions), Git, JIRA"),
        ]
    }
}
