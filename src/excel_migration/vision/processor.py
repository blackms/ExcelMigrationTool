"""Image processing capabilities for Excel sheet analysis."""
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import cv2
import numpy as np
import pytesseract
from PIL import Image
import torch
from transformers import AutoProcessor, AutoModelForVision2Seq
from loguru import logger

from ..core.interfaces import ImageProcessor

class SheetImageProcessor(ImageProcessor):
    """Process Excel sheet images for data extraction."""
    
    def __init__(self):
        """Initialize the image processor with required models."""
        self.vision_processor = AutoProcessor.from_pretrained(
            "microsoft/git-base-coco"
        )
        self.vision_model = AutoModelForVision2Seq.from_pretrained(
            "microsoft/git-base-coco"
        )
        logger.debug("Initialized sheet image processor")

    async def process_image(self, image_path: Path) -> Dict[str, Any]:
        """Process an image and extract information."""
        try:
            # Load and preprocess image
            image = cv2.imread(str(image_path))
            if image is None:
                raise ValueError(f"Failed to load image: {image_path}")

            # Extract various features
            result = {
                "table_structure": await self._detect_table_structure(image),
                "text_content": await self._extract_text(image),
                "cell_boundaries": await self._detect_cells(image),
                "visual_analysis": await self._analyze_visual_elements(image),
                "layout_analysis": await self._analyze_layout(image)
            }

            logger.info(f"Processed image: {image_path}")
            return result

        except Exception as e:
            logger.exception(f"Image processing failed: {str(e)}")
            return {"error": str(e)}

    async def extract_table(self, image: np.ndarray) -> Optional[List[List[str]]]:
        """Extract table data from an image."""
        try:
            # Detect table structure
            table_structure = await self._detect_table_structure(image)
            if not table_structure:
                return None

            # Extract cells
            cells = await self._detect_cells(image)
            
            # Extract text from each cell
            table_data = []
            for row in cells:
                row_data = []
                for cell in row:
                    text = await self._extract_cell_text(image, cell)
                    row_data.append(text)
                table_data.append(row_data)

            return table_data

        except Exception as e:
            logger.exception(f"Table extraction failed: {str(e)}")
            return None

    async def _detect_table_structure(self, image: np.ndarray) -> Dict[str, Any]:
        """Detect table structure in the image."""
        try:
            # Convert to grayscale
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # Apply adaptive thresholding
            thresh = cv2.adaptiveThreshold(
                gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
                cv2.THRESH_BINARY_INV, 11, 2
            )
            
            # Detect lines
            horizontal = await self._detect_lines(thresh, True)
            vertical = await self._detect_lines(thresh, False)
            
            # Combine lines
            table_structure = cv2.add(horizontal, vertical)
            
            # Find contours
            contours, _ = cv2.findContours(
                table_structure, 
                cv2.RETR_TREE, 
                cv2.CHAIN_APPROX_SIMPLE
            )
            
            return {
                "structure": table_structure,
                "contours": contours,
                "horizontal_lines": horizontal,
                "vertical_lines": vertical
            }

        except Exception as e:
            logger.exception(f"Table structure detection failed: {str(e)}")
            return {}

    async def _detect_lines(self, image: np.ndarray, horizontal: bool) -> np.ndarray:
        """Detect lines in the image."""
        # Get image dimensions
        (h, w) = image.shape
        
        # Create structure element
        if horizontal:
            structure = cv2.getStructuringElement(cv2.MORPH_RECT, (w // 30, 1))
        else:
            structure = cv2.getStructuringElement(cv2.MORPH_RECT, (1, h // 30))
        
        # Apply morphology
        return cv2.morphologyEx(image, cv2.MORPH_OPEN, structure)

    async def _detect_cells(self, image: np.ndarray) -> List[List[Tuple[int, int, int, int]]]:
        """Detect cell boundaries in the image."""
        try:
            # Get table structure
            structure = await self._detect_table_structure(image)
            if not structure:
                return []
            
            # Find cell contours
            contours = structure["contours"]
            
            # Filter and sort contours
            cells = []
            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                if w * h > 100:  # Filter small contours
                    cells.append((x, y, w, h))
            
            # Sort cells by position
            cells.sort(key=lambda c: (c[1], c[0]))  # Sort by y, then x
            
            # Group cells into rows
            rows = []
            current_row = []
            current_y = cells[0][1] if cells else 0
            
            for cell in cells:
                if abs(cell[1] - current_y) > 10:  # New row
                    if current_row:
                        rows.append(current_row)
                    current_row = [cell]
                    current_y = cell[1]
                else:
                    current_row.append(cell)
            
            if current_row:
                rows.append(current_row)
            
            return rows

        except Exception as e:
            logger.exception(f"Cell detection failed: {str(e)}")
            return []

    async def _extract_text(self, image: np.ndarray) -> str:
        """Extract text from the image using OCR."""
        try:
            # Convert to PIL Image
            pil_image = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
            
            # Extract text
            text = pytesseract.image_to_string(pil_image)
            
            return text.strip()

        except Exception as e:
            logger.exception(f"Text extraction failed: {str(e)}")
            return ""

    async def _extract_cell_text(self, image: np.ndarray, 
                               cell: Tuple[int, int, int, int]) -> str:
        """Extract text from a specific cell."""
        try:
            x, y, w, h = cell
            cell_image = image[y:y+h, x:x+w]
            return await self._extract_text(cell_image)

        except Exception as e:
            logger.exception(f"Cell text extraction failed: {str(e)}")
            return ""

    async def _analyze_visual_elements(self, image: np.ndarray) -> Dict[str, Any]:
        """Analyze visual elements using vision model."""
        try:
            # Convert to PIL Image
            pil_image = Image.fromarray(cv2.cvtColor(image, cv2.COLOR_BGR2RGB))
            
            # Process image with vision model
            inputs = self.vision_processor(
                images=pil_image, 
                return_tensors="pt"
            )
            
            outputs = self.vision_model.generate(
                pixel_values=inputs.pixel_values,
                max_length=50
            )
            
            # Decode prediction
            prediction = self.vision_processor.decode(
                outputs[0], 
                skip_special_tokens=True
            )
            
            return {
                "description": prediction,
                "confidence": float(outputs.sequences_scores[0])
                if hasattr(outputs, "sequences_scores")
                else None
            }

        except Exception as e:
            logger.exception(f"Visual analysis failed: {str(e)}")
            return {}

    async def _analyze_layout(self, image: np.ndarray) -> Dict[str, Any]:
        """Analyze the layout structure of the sheet."""
        try:
            # Convert to grayscale
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # Apply thresholding
            _, thresh = cv2.threshold(
                gray, 0, 255, 
                cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU
            )
            
            # Find contours
            contours, hierarchy = cv2.findContours(
                thresh, 
                cv2.RETR_TREE, 
                cv2.CHAIN_APPROX_SIMPLE
            )
            
            # Analyze layout structure
            layout = {
                "regions": [],
                "hierarchy": hierarchy.tolist() if hierarchy is not None else None
            }
            
            for contour in contours:
                x, y, w, h = cv2.boundingRect(contour)
                area = cv2.contourArea(contour)
                if area > 100:  # Filter small regions
                    layout["regions"].append({
                        "bounds": (x, y, w, h),
                        "area": area,
                        "type": self._classify_region(w, h, area)
                    })
            
            return layout

        except Exception as e:
            logger.exception(f"Layout analysis failed: {str(e)}")
            return {}

    def _classify_region(self, width: int, height: int, area: float) -> str:
        """Classify a region based on its properties."""
        aspect_ratio = width / height if height != 0 else 0
        
        if aspect_ratio > 3:
            return "header"
        elif aspect_ratio < 0.3:
            return "column"
        elif area > 10000:
            return "table"
        else:
            return "cell"