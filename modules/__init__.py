# -*- coding: utf-8 -*-
"""
NexaScout Core Modules
- enrichment: Web scraping, contact info, email and social media finder
- map_generator: Interactive Folium GIS map generation
- pitch_generator: AI-powered personalized investment proposal generation (DOCX)
"""

from .enrichment import clean_phone_for_whatsapp, generate_whatsapp_url, enrich_lead_contacts
from .map_generator import generate_interactive_map
from .pitch_generator import generate_pitch_document

__all__ = [
    "clean_phone_for_whatsapp",
    "generate_whatsapp_url",
    "enrich_lead_contacts",
    "generate_interactive_map",
    "generate_pitch_document"
]
