# KherveFitting - XPS Data Analysis Software
# Copyright (C) 2024-2026 Gwilherm Kerherve <g.kerherve@ic.ac.uk>
#
# KherveFitting is dual-licensed:
#   - GNU GPL v3.0 (see LICENSE-GPL.txt) for open-source use
#   - Commercial Licence (see LICENSE-COMMERCIAL.txt) for proprietary use
# SPDX-License-Identifier: GPL-3.0-only OR LicenseRef-KherveFitting-Commercial

"""
BCF_Reader.py
Reader for Bruker BCF (Binary Container Format) files.

BCF files are ZIP archives containing:
- XML metadata files
- Binary spectrum/map data
- SEM images

This module provides BCF reading without HyperSpy dependency.

Author: KherveFitting Project
License: BSD-3-Clause
"""

import zipfile
import xml.etree.ElementTree as ET
import numpy as np
import struct
import io
import os
from typing import Dict, List, Tuple, Optional, Any
from dataclasses import dataclass, field


@dataclass
class BCFSpectrum:
    """Container for a single EDX spectrum"""
    data: np.ndarray
    energy: np.ndarray
    metadata: Dict[str, Any] = field(default_factory=dict)
    elements: List[str] = field(default_factory=list)

    @property
    def shape(self):
        return self.data.shape


@dataclass
class BCFMap:
    """Container for EDX element map or hyperspectral data"""
    data: np.ndarray  # Can be 2D (element map) or 3D (hyperspectral)
    energy: Optional[np.ndarray] = None
    metadata: Dict[str, Any] = field(default_factory=dict)
    elements: List[str] = field(default_factory=list)
    pixel_size_x: float = 1.0  # µm
    pixel_size_y: float = 1.0  # µm

    @property
    def shape(self):
        return self.data.shape

    @property
    def is_hyperspectral(self):
        return self.data.ndim == 3


@dataclass
class BCFImage:
    """Container for SEM image"""
    data: np.ndarray
    metadata: Dict[str, Any] = field(default_factory=dict)
    pixel_size_x: float = 1.0
    pixel_size_y: float = 1.0
    detector: str = ''


@dataclass
class BCFData:
    """Complete BCF file data"""
    spectra: List[BCFSpectrum] = field(default_factory=list)
    maps: List[BCFMap] = field(default_factory=list)
    images: List[BCFImage] = field(default_factory=list)
    metadata: Dict[str, Any] = field(default_factory=dict)
    filename: str = ''

    @property
    def has_spectra(self):
        return len(self.spectra) > 0

    @property
    def has_maps(self):
        return len(self.maps) > 0

    @property
    def has_images(self):
        return len(self.images) > 0


class BCFReader:
    """
    Reader for Bruker BCF files.

    BCF files are ZIP archives with XML metadata and binary data.
    """

    def __init__(self, filepath: str):
        """
        Initialize BCF reader.

        Parameters
        ----------
        filepath : str
            Path to BCF file
        """
        self.filepath = filepath
        self.zipfile = None
        self._header_xml = None
        self._data_xml = None
        self._metadata = {}

    def __enter__(self):
        self.open()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
        return False

    def open(self):
        """Open the BCF file"""
        if not os.path.exists(self.filepath):
            raise FileNotFoundError(f"BCF file not found: {self.filepath}")
        self.zipfile = zipfile.ZipFile(self.filepath, 'r')

    def close(self):
        """Close the BCF file"""
        if self.zipfile:
            self.zipfile.close()
            self.zipfile = None

    def list_contents(self) -> List[str]:
        """List files in the BCF archive"""
        if not self.zipfile:
            self.open()
        return self.zipfile.namelist()

    def read(self) -> BCFData:
        """
        Read all data from BCF file.

        Returns
        -------
        BCFData
            Container with spectra, maps, and images
        """
        if not self.zipfile:
            self.open()

        result = BCFData(filename=os.path.basename(self.filepath))

        # Parse XML metadata
        self._parse_header()
        result.metadata = self._metadata.copy()

        # Read spectra
        spectra = self._read_spectra()
        result.spectra.extend(spectra)

        # Read maps
        maps = self._read_maps()
        result.maps.extend(maps)

        # Read images
        images = self._read_images()
        result.images.extend(images)

        return result

    def _parse_header(self):
        """Parse the header XML file"""
        try:
            # Find header file
            names = self.zipfile.namelist()
            header_files = [n for n in names if 'header' in n.lower() and n.endswith('.xml')]

            if not header_files:
                # Try common patterns
                header_files = [n for n in names if n.endswith('.xml')]

            if header_files:
                with self.zipfile.open(header_files[0]) as f:
                    content = f.read()
                    try:
                        self._header_xml = ET.fromstring(content)
                        self._extract_metadata_from_xml(self._header_xml)
                    except ET.ParseError:
                        pass
        except Exception as e:
            print(f"Warning: Could not parse header: {e}")

    def _extract_metadata_from_xml(self, root: ET.Element):
        """Extract metadata from XML element tree"""

        def recursive_extract(element, prefix=''):
            for child in element:
                tag = child.tag.split('}')[-1]  # Remove namespace
                key = f"{prefix}{tag}" if prefix else tag

                if child.text and child.text.strip():
                    self._metadata[key] = child.text.strip()

                # Extract attributes
                for attr, value in child.attrib.items():
                    self._metadata[f"{key}_{attr}"] = value

                # Recurse
                recursive_extract(child, f"{key}/")

        recursive_extract(root)

    def _read_spectra(self) -> List[BCFSpectrum]:
        """Read spectrum data from BCF"""
        spectra = []

        try:
            names = self.zipfile.namelist()

            # Look for spectrum files
            spectrum_files = [n for n in names if 'spectrum' in n.lower()
                              or n.endswith('.spx') or 'eds' in n.lower()]

            for spec_file in spectrum_files:
                try:
                    spectrum = self._read_single_spectrum(spec_file)
                    if spectrum is not None:
                        spectra.append(spectrum)
                except Exception as e:
                    print(f"Warning: Could not read spectrum {spec_file}: {e}")

        except Exception as e:
            print(f"Warning: Error reading spectra: {e}")

        return spectra

    def _read_single_spectrum(self, filename: str) -> Optional[BCFSpectrum]:
        """Read a single spectrum file"""
        with self.zipfile.open(filename) as f:
            content = f.read()

        # Try to parse as XML (common for SPX files)
        try:
            root = ET.fromstring(content)
            return self._parse_spx_spectrum(root)
        except ET.ParseError:
            pass

        # Try as binary data
        try:
            return self._parse_binary_spectrum(content)
        except Exception:
            pass

        return None

    def _parse_spx_spectrum(self, root: ET.Element) -> Optional[BCFSpectrum]:
        """Parse SPX format spectrum (XML)"""
        metadata = {}
        data = None
        energy = None

        # Extract spectrum parameters
        for elem in root.iter():
            tag = elem.tag.split('}')[-1]

            if tag == 'ChannelCount':
                metadata['channels'] = int(elem.text)
            elif tag == 'CalibAbs':
                metadata['offset'] = float(elem.text)
            elif tag == 'CalibLin':
                metadata['scale'] = float(elem.text)
            elif tag == 'Channels' or tag == 'Data':
                # Spectrum data
                if elem.text:
                    values = elem.text.strip().split(',')
                    data = np.array([float(v) for v in values if v.strip()])

        if data is None:
            return None

        # Generate energy axis
        n_channels = len(data)
        offset = metadata.get('offset', 0)
        scale = metadata.get('scale', 0.01)  # Default 10 eV/channel
        energy = np.arange(n_channels) * scale + offset

        return BCFSpectrum(data=data, energy=energy, metadata=metadata)

    def _parse_binary_spectrum(self, content: bytes) -> Optional[BCFSpectrum]:
        """Parse binary spectrum data"""
        # Try common binary formats

        # Try as 32-bit float
        try:
            data = np.frombuffer(content, dtype=np.float32)
            if len(data) > 100 and len(data) < 50000:
                energy = np.arange(len(data)) * 0.01  # Default 10 eV/channel
                return BCFSpectrum(data=data, energy=energy)
        except Exception:
            pass

        # Try as 32-bit int
        try:
            data = np.frombuffer(content, dtype=np.int32)
            if len(data) > 100 and len(data) < 50000:
                energy = np.arange(len(data)) * 0.01
                return BCFSpectrum(data=data.astype(float), energy=energy)
        except Exception:
            pass

        return None

    def _read_maps(self) -> List[BCFMap]:
        """Read element maps from BCF"""
        maps = []

        try:
            names = self.zipfile.namelist()

            # Look for map/image data files
            map_files = [n for n in names if 'map' in n.lower()
                         or 'hypermap' in n.lower()
                         or n.endswith('.bcf')]

            for map_file in map_files:
                try:
                    map_data = self._read_single_map(map_file)
                    if map_data is not None:
                        maps.append(map_data)
                except Exception as e:
                    print(f"Warning: Could not read map {map_file}: {e}")

        except Exception as e:
            print(f"Warning: Error reading maps: {e}")

        # Also try to read hyperspectral data directly
        try:
            hyper = self._read_hyperspectral()
            if hyper is not None:
                maps.append(hyper)
        except Exception as e:
            print(f"Warning: Could not read hyperspectral data: {e}")

        return maps

    def _read_single_map(self, filename: str) -> Optional[BCFMap]:
        """Read a single map file"""
        with self.zipfile.open(filename) as f:
            content = f.read()

        # Try to determine dimensions from metadata
        width = self._metadata.get('ImageWidth', self._metadata.get('Width', 256))
        height = self._metadata.get('ImageHeight', self._metadata.get('Height', 256))

        try:
            width = int(width)
            height = int(height)
        except (ValueError, TypeError):
            width, height = 256, 256

        # Try to read as image
        try:
            data = np.frombuffer(content, dtype=np.uint16)
            if len(data) == width * height:
                data = data.reshape((height, width))
                return BCFMap(data=data.astype(float))
        except Exception:
            pass

        try:
            data = np.frombuffer(content, dtype=np.float32)
            if len(data) == width * height:
                data = data.reshape((height, width))
                return BCFMap(data=data)
        except Exception:
            pass

        return None

    def _read_hyperspectral(self) -> Optional[BCFMap]:
        """Read hyperspectral datacube"""
        names = self.zipfile.namelist()

        # Look for raw data files
        raw_files = [n for n in names if n.endswith('.raw')
                     or 'datacube' in n.lower()
                     or 'hyperspectral' in n.lower()]

        if not raw_files:
            return None

        # Get dimensions
        width = int(self._metadata.get('Width', self._metadata.get('ImageWidth', 128)))
        height = int(self._metadata.get('Height', self._metadata.get('ImageHeight', 128)))
        channels = int(self._metadata.get('Channels', self._metadata.get('ChannelCount', 2048)))

        for raw_file in raw_files:
            try:
                with self.zipfile.open(raw_file) as f:
                    content = f.read()

                # Try different data types
                for dtype in [np.uint16, np.uint32, np.float32]:
                    data = np.frombuffer(content, dtype=dtype)

                    # Check if size matches expected cube dimensions
                    expected_size = width * height * channels
                    if len(data) == expected_size:
                        data = data.reshape((height, width, channels))

                        # Generate energy axis
                        offset = float(self._metadata.get('CalibAbs', 0))
                        scale = float(self._metadata.get('CalibLin', 0.01))
                        energy = np.arange(channels) * scale + offset

                        return BCFMap(
                            data=data.astype(float),
                            energy=energy,
                            metadata={'source': raw_file}
                        )
            except Exception:
                continue

        return None

    def _read_images(self) -> List[BCFImage]:
        """Read SEM images from BCF"""
        images = []

        try:
            names = self.zipfile.namelist()

            # Look for image files
            image_exts = ('.tif', '.tiff', '.png', '.jpg', '.jpeg', '.bmp')
            image_files = [n for n in names if n.lower().endswith(image_exts)
                           or 'image' in n.lower()]

            for img_file in image_files:
                try:
                    with self.zipfile.open(img_file) as f:
                        content = f.read()

                    # Try to read with PIL/numpy
                    img = self._decode_image(content)
                    if img is not None:
                        images.append(BCFImage(data=img, metadata={'source': img_file}))
                except Exception as e:
                    print(f"Warning: Could not read image {img_file}: {e}")

        except Exception as e:
            print(f"Warning: Error reading images: {e}")

        return images

    def _decode_image(self, content: bytes) -> Optional[np.ndarray]:
        """Decode image from bytes"""
        try:
            from PIL import Image
            img = Image.open(io.BytesIO(content))
            return np.array(img)
        except ImportError:
            pass
        except Exception:
            pass

        # Try raw decode
        try:
            # Look for TIFF header
            if content[:2] in (b'II', b'MM'):
                # TIFF file - try basic decode
                # This is a simplified TIFF reader
                pass
        except Exception:
            pass

        return None


def load_bcf(filepath: str) -> BCFData:
    """
    Load a Bruker BCF file.

    Parameters
    ----------
    filepath : str
        Path to BCF file

    Returns
    -------
    BCFData
        Container with spectra, maps, and images
    """
    with BCFReader(filepath) as reader:
        return reader.read()


def load_bcf_spectrum(filepath: str) -> Optional[BCFSpectrum]:
    """
    Load the first/main spectrum from a BCF file.

    Parameters
    ----------
    filepath : str
        Path to BCF file

    Returns
    -------
    BCFSpectrum or None
        Spectrum data or None if not found
    """
    data = load_bcf(filepath)
    if data.spectra:
        return data.spectra[0]
    return None


def load_bcf_map(filepath: str) -> Optional[BCFMap]:
    """
    Load the first/main map from a BCF file.

    Parameters
    ----------
    filepath : str
        Path to BCF file

    Returns
    -------
    BCFMap or None
        Map data or None if not found
    """
    data = load_bcf(filepath)
    if data.maps:
        return data.maps[0]
    return None


# Aliases for compatibility
read_bcf = load_bcf

__all__ = [
    'BCFReader',
    'BCFData',
    'BCFSpectrum',
    'BCFMap',
    'BCFImage',
    'load_bcf',
    'load_bcf_spectrum',
    'load_bcf_map',
    'read_bcf',
]