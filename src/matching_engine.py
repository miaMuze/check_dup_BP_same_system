"""
BP Duplicate Checker - Fuzzy Matching Engine
=============================================
This module contains the core fuzzy matching logic for detecting
potential duplicate Business Partners based on name similarity.

Key features:
- Text normalization (lowercase, trim spaces, remove punctuation)
- Configurable ignore word list
- RapidFuzz-based similarity scoring
- Returns top N most similar matches for each BP
"""

import re
import string
from typing import List, Dict, Tuple, Optional
from dataclasses import dataclass
from rapidfuzz import fuzz, process


@dataclass
class BPRecord:
    """
    Represents a single Business Partner record.

    Attributes:
        bp_number: Unique identifier for the Business Partner
        name1: Primary name field
        name2: Secondary name field (can be empty)
        combined_name: Normalized combination of name1 and name2
    """
    bp_number: str
    name1: str
    name2: str
    combined_name: str = ""

    def __post_init__(self):
        """Combine name1 and name2 after initialization."""
        self.combined_name = f"{self.name1} {self.name2}".strip()


@dataclass
class MatchResult:
    """
    Represents a single match result between two BPs.

    Attributes:
        source_bp: The BP being compared
        match_bp: The potentially duplicate BP
        similarity_score: Score from 0-100 indicating similarity
    """
    source_bp: BPRecord
    match_bp: BPRecord
    similarity_score: float


class TextNormalizer:
    """
    Handles text normalization for consistent comparison.

    Normalization steps:
    1. Convert to lowercase
    2. Remove punctuation
    3. Remove extra whitespace
    4. Remove ignored words (configurable)
    """

    # Default words to ignore during comparison
    DEFAULT_IGNORE_WORDS = {
        'mrs', 'ms', 'mr', 'dr', 'prof',
        'company', 'co', 'ltd', 'llc', 'inc', 'corp', 'corporation',
        'limited', 'plc', 'gmbh', 'ag', 'sa', 'srl',
        'the', 'and', '&'
    }

    def __init__(self, ignore_words: Optional[List[str]] = None):
        """
        Initialize the normalizer with optional custom ignore words.

        Args:
            ignore_words: List of words to exclude from comparison.
                         If None, uses default list.
        """
        if ignore_words is not None:
            # Convert user-provided words to lowercase set
            self.ignore_words = {word.lower().strip() for word in ignore_words if word.strip()}
        else:
            self.ignore_words = self.DEFAULT_IGNORE_WORDS.copy()

    def normalize(self, text: str) -> str:
        """
        Normalize text for comparison.

        Args:
            text: Raw text to normalize

        Returns:
            Normalized text string
        """
        if not text:
            return ""

        # Step 1: Convert to lowercase
        normalized = text.lower()

        # Step 2: Remove punctuation (replace with spaces to preserve word boundaries)
        # Keep alphanumeric and spaces only
        normalized = re.sub(r'[^\w\s]', ' ', normalized)

        # Step 3: Split into words
        words = normalized.split()

        # Step 4: Remove ignored words
        words = [word for word in words if word not in self.ignore_words]

        # Step 5: Rejoin and collapse multiple spaces
        normalized = ' '.join(words)

        return normalized.strip()


class FuzzyMatcher:
    """
    Main fuzzy matching engine for BP duplicate detection.

    Uses RapidFuzz library for efficient string matching with
    multiple scoring algorithms combined for better accuracy.
    """

    def __init__(self, ignore_words: Optional[List[str]] = None):
        """
        Initialize the matcher with optional ignore words.

        Args:
            ignore_words: List of words to exclude from comparison
        """
        self.normalizer = TextNormalizer(ignore_words)
        self.records: List[BPRecord] = []
        self.normalized_names: Dict[str, str] = {}  # bp_number -> normalized name

    def load_records(self, data: List[Dict[str, str]]) -> int:
        """
        Load BP records from a list of dictionaries.

        Args:
            data: List of dicts with keys 'BP_Number', 'Name1', 'Name2'

        Returns:
            Number of records loaded
        """
        self.records = []
        self.normalized_names = {}

        for row in data:
            bp_number = str(row.get('BP_Number', '')).strip()
            name1 = str(row.get('Name1', '')).strip()
            name2 = str(row.get('Name2', '')).strip()

            if not bp_number:
                continue  # Skip records without BP number

            record = BPRecord(
                bp_number=bp_number,
                name1=name1,
                name2=name2
            )
            self.records.append(record)

            # Pre-compute normalized names for efficiency
            self.normalized_names[bp_number] = self.normalizer.normalize(
                record.combined_name
            )

        return len(self.records)

    def calculate_similarity(self, name1: str, name2: str) -> float:
        """
        Calculate similarity score between two normalized names.

        Uses a weighted combination of different fuzzy matching algorithms:
        - Token Sort Ratio: Good for names with words in different order
        - Token Set Ratio: Good for names with different word counts
        - Ratio: Standard Levenshtein-based similarity

        Args:
            name1: First normalized name
            name2: Second normalized name

        Returns:
            Similarity score from 0 to 100
        """
        if not name1 or not name2:
            return 0.0

        # Calculate different similarity metrics
        token_sort = fuzz.token_sort_ratio(name1, name2)
        token_set = fuzz.token_set_ratio(name1, name2)
        simple_ratio = fuzz.ratio(name1, name2)

        # Weighted average (token-based scores are more important for names)
        # Token Sort: 40% - handles word order differences
        # Token Set: 40% - handles partial matches
        # Simple Ratio: 20% - exact character matching
        weighted_score = (
            token_sort * 0.4 +
            token_set * 0.4 +
            simple_ratio * 0.2
        )

        return round(weighted_score, 2)

    def find_matches(
        self,
        top_n: int = 3,
        min_score: float = 50.0,
        progress_callback=None
    ) -> Dict[str, List[MatchResult]]:
        """
        Find potential duplicate matches for all BP records.

        Args:
            top_n: Number of top matches to return for each BP
            min_score: Minimum similarity score to consider as a match
            progress_callback: Optional callback function(current, total) for progress updates

        Returns:
            Dictionary mapping BP_Number to list of MatchResults
        """
        results: Dict[str, List[MatchResult]] = {}
        total = len(self.records)

        for idx, source in enumerate(self.records):
            source_normalized = self.normalized_names.get(source.bp_number, '')
            matches: List[MatchResult] = []

            # Compare with all other records
            for target in self.records:
                # Skip self-comparison
                if source.bp_number == target.bp_number:
                    continue

                target_normalized = self.normalized_names.get(target.bp_number, '')

                # Calculate similarity
                score = self.calculate_similarity(source_normalized, target_normalized)

                # Only include if above minimum threshold
                if score >= min_score:
                    matches.append(MatchResult(
                        source_bp=source,
                        match_bp=target,
                        similarity_score=score
                    ))

            # Sort by similarity score (descending) and take top N
            matches.sort(key=lambda x: x.similarity_score, reverse=True)
            results[source.bp_number] = matches[:top_n]

            # Report progress
            if progress_callback:
                progress_callback(idx + 1, total)

        return results

    def get_summary_stats(self, results: Dict[str, List[MatchResult]]) -> Dict:
        """
        Generate summary statistics for the matching results.

        Args:
            results: Matching results from find_matches()

        Returns:
            Dictionary with summary statistics
        """
        total_records = len(self.records)
        records_with_matches = sum(1 for matches in results.values() if matches)

        all_scores = []
        high_confidence_count = 0  # Score >= 80
        medium_confidence_count = 0  # Score 60-79
        low_confidence_count = 0  # Score < 60

        for matches in results.values():
            for match in matches:
                score = match.similarity_score
                all_scores.append(score)

                if score >= 80:
                    high_confidence_count += 1
                elif score >= 60:
                    medium_confidence_count += 1
                else:
                    low_confidence_count += 1

        avg_score = sum(all_scores) / len(all_scores) if all_scores else 0

        return {
            'total_records': total_records,
            'records_with_matches': records_with_matches,
            'total_matches': len(all_scores),
            'average_score': round(avg_score, 2),
            'high_confidence': high_confidence_count,
            'medium_confidence': medium_confidence_count,
            'low_confidence': low_confidence_count
        }
