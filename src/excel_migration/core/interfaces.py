"""Core interfaces for the Excel migration framework."""
from abc import ABC, abstractmethod
from typing import Any, Dict, List, Optional, Protocol
from pathlib import Path

class Task(Protocol):
    """Protocol for migration tasks."""
    source_file: Path
    target_file: Path
    task_type: str
    description: str
    context: Dict[str, Any]

class RuleGenerator(Protocol):
    """Protocol for rule generation."""
    async def generate_rules(self, source: Path, target: Path) -> List[Dict[str, Any]]:
        """Generate rules by analyzing source and target files."""
        ...

class SheetAnalyzer(Protocol):
    """Protocol for sheet analysis."""
    async def analyze_sheet(self, sheet_path: Path) -> Dict[str, Any]:
        """Analyze a sheet and return its structure and content."""
        ...

class DataExtractor(Protocol):
    """Protocol for data extraction."""
    async def extract_data(self, image_path: Path) -> Dict[str, Any]:
        """Extract data from sheet images."""
        ...

class TaskHandler(ABC):
    """Abstract base class for task handlers."""
    
    @abstractmethod
    async def can_handle(self, task: Task) -> bool:
        """Check if this handler can process the task."""
        pass
    
    @abstractmethod
    async def handle(self, task: Task) -> bool:
        """Handle the task."""
        pass

class TaskProcessor(ABC):
    """Abstract base class for task processors."""
    
    @abstractmethod
    async def process(self, task: Task) -> bool:
        """Process a task."""
        pass
    
    @abstractmethod
    async def validate(self, task: Task) -> bool:
        """Validate task requirements."""
        pass

class RuleExecutor(ABC):
    """Abstract base class for rule execution."""
    
    @abstractmethod
    async def execute(self, rule: Dict[str, Any], context: Dict[str, Any]) -> Any:
        """Execute a single rule."""
        pass
    
    @abstractmethod
    async def validate_rule(self, rule: Dict[str, Any]) -> bool:
        """Validate a rule's structure and requirements."""
        pass

class ImageProcessor(ABC):
    """Abstract base class for image processing."""
    
    @abstractmethod
    async def process_image(self, image_path: Path) -> Dict[str, Any]:
        """Process an image and extract information."""
        pass
    
    @abstractmethod
    async def extract_table(self, image: Any) -> Optional[List[List[str]]]:
        """Extract table data from an image."""
        pass

class LLMProvider(ABC):
    """Abstract base class for LLM providers."""
    
    @abstractmethod
    async def analyze_task(self, task: Task) -> Dict[str, Any]:
        """Analyze a task and provide insights."""
        pass
    
    @abstractmethod
    async def generate_rules(self, source_data: Dict[str, Any], 
                           target_data: Dict[str, Any]) -> List[Dict[str, Any]]:
        """Generate migration rules based on source and target data."""
        pass
    
    @abstractmethod
    async def validate_transformation(self, source: Any, target: Any, 
                                   rules: List[Dict[str, Any]]) -> bool:
        """Validate transformation results."""
        pass

class Logger(Protocol):
    """Protocol for logging interface."""
    def debug(self, message: str, **kwargs: Any) -> None: ...
    def info(self, message: str, **kwargs: Any) -> None: ...
    def warning(self, message: str, **kwargs: Any) -> None: ...
    def error(self, message: str, **kwargs: Any) -> None: ...
    def exception(self, message: str, **kwargs: Any) -> None: ...

class ConfigProvider(Protocol):
    """Protocol for configuration management."""
    def get_config(self, key: str) -> Any: ...
    def set_config(self, key: str, value: Any) -> None: ...
    def load_config(self, path: Path) -> None: ...
    def save_config(self, path: Path) -> None: ...

class CacheProvider(Protocol):
    """Protocol for caching."""
    async def get(self, key: str) -> Optional[Any]: ...
    async def set(self, key: str, value: Any, ttl: Optional[int] = None) -> None: ...
    async def delete(self, key: str) -> None: ...
    async def clear(self) -> None: ...

class EventEmitter(Protocol):
    """Protocol for event handling."""
    def emit(self, event: str, data: Any) -> None: ...
    def on(self, event: str, callback: callable) -> None: ...
    def off(self, event: str, callback: callable) -> None: ...

class MetricsCollector(Protocol):
    """Protocol for metrics collection."""
    def record_metric(self, name: str, value: float, tags: Optional[Dict[str, str]] = None) -> None: ...
    def increment_counter(self, name: str, tags: Optional[Dict[str, str]] = None) -> None: ...
    def start_timer(self, name: str) -> None: ...
    def stop_timer(self, name: str) -> float: ...