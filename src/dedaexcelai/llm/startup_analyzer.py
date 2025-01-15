import os
from openai import OpenAI
from typing import Optional
from ..logger import get_logger

logger = get_logger()

class StartupDaysAnalyzer:
    def __init__(self):
        self.client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))
        
    def analyze_startup_days(self, description: str) -> Optional[int]:
        """
        Analyze the description from column H to determine the number of startup days.
        
        Args:
            description: The text from column H describing the service/product
            
        Returns:
            Optional[int]: The number of startup days if determinable, None otherwise
        """
        try:
            # Construct a prompt that focuses on extracting startup days
            prompt = f"""
            Analizza il seguente testo che descrive un prodotto/servizio e determina il numero di giorni necessari per lo startup:

            {description}

            Considera:
            - Cerca riferimenti espliciti a giorni di setup/startup/installazione
            - Considera la complessità del prodotto/servizio descritto
            - Se non ci sono riferimenti espliciti ma il contesto suggerisce un tempo di setup, fai una stima ragionevole
            - Se non ci sono sufficienti informazioni, restituisci None

            Rispondi SOLO con il numero di giorni (un numero intero) o 'None' se non determinabile.
            """

            # Get completion from GPT
            completion = self.client.chat.completions.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "Sei un esperto di analisi dei tempi di setup e startup di prodotti e servizi IT."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.2,  # Lower temperature for more consistent numerical outputs
                max_tokens=10     # We only need a number or 'None'
            )
            
            # Extract the response
            response = completion.choices[0].message.content.strip()
            
            # Convert response to integer if possible
            if response.lower() == 'none':
                return None
            try:
                return int(response)
            except ValueError:
                logger.warning(f"Unexpected response format from LLM: {response}")
                return None
                
        except Exception as e:
            logger.error(f"Error analyzing startup days: {str(e)}")
            return None
