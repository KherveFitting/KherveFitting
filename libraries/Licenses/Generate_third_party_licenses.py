# generate_third_party_licenses.py
import importlib.metadata
import sys
from pathlib import Path


def get_package_license_info(package_name):
    """Get license information for a package"""
    try:
        metadata = importlib.metadata.metadata(package_name)

        # Try different metadata fields for license info
        license_info = (
            metadata.get('License') or
            metadata.get('License-Expression') or
            metadata.get('Classifier', '').split('License ::')[-1].strip() if 'License ::' in str(
                metadata.get('Classifier', '')) else
            'License not specified'
        )

        # Clean up license info
        if isinstance(license_info, list):
            license_info = ', '.join(license_info)

        homepage = (
                metadata.get('Home-page') or
                metadata.get('Project-URL') or
                'Homepage not specified'
        )

        if isinstance(homepage, list):
            homepage = homepage[0] if homepage else 'Homepage not specified'

        return {
            'version': metadata.get('Version', 'Unknown'),
            'license': license_info,
            'homepage': homepage,
            'summary': metadata.get('Summary', 'No description available')
        }

    except importlib.metadata.PackageNotFoundError:
        return None
    except Exception as e:
        return {'error': str(e)}


def generate_third_party_licenses():
    """Generate THIRD_PARTY_LICENSES file with better error handling"""

    # Core dependencies for your XPS application
    required_packages = [
        'wxpython', 'matplotlib', 'numpy', 'pandas', 'openpyxl',
        'scipy', 'pillow', 'lmfit', 'setuptools', 'wheel'
    ]

    licenses_content = [
        "THIRD-PARTY SOFTWARE LICENSES",
        "=" * 50,
        "",
        "This software uses the following third-party libraries:",
        ""
    ]

    for package_name in required_packages:
        info = get_package_license_info(package_name)

        if info is None:
            licenses_content.extend([
                f"{package_name.upper()}",
                "-" * len(package_name),
                "Status: Not installed or not found",
                ""
            ])
            continue

        if 'error' in info:
            licenses_content.extend([
                f"{package_name.upper()}",
                "-" * len(package_name),
                f"Error retrieving info: {info['error']}",
                ""
            ])
            continue

        # Add package info
        header = f"{package_name.upper()} v{info['version']}"
        licenses_content.extend([
            header,
            "-" * len(header),
            f"License: {info['license']}",
            f"Homepage: {info['homepage']}",
            f"Description: {info['summary']}",
            ""
        ])

    # Add common license texts for major licenses
    licenses_content.extend([
        "",
        "COMMON LICENSE TEXTS",
        "=" * 20,
        "",
        "BSD-3-Clause License:",
        "Redistribution and use in source and binary forms, with or without modification,",
        "are permitted provided that conditions are met (see individual package licenses).",
        "",
        "MIT License:",
        "Permission is hereby granted, free of charge, to use, copy, modify, merge, publish,",
        "distribute, sublicense, and/or sell copies (see individual package licenses).",
        "",
        "For complete license texts, visit the individual package homepages listed above.",
    ])

    # Write to file
    output_path = Path('../../THIRD_PARTY_LICENSES.txt')
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(licenses_content))
        print(f"Successfully created {output_path}")

        # Also create a simple version for embedding
        create_simple_licenses_file(required_packages)

    except Exception as e:
        print(f"Error writing license file: {e}")


def create_simple_licenses_file(packages):
    """Create a simplified version for embedding in applications"""
    simple_content = [
        "Third-Party Software Licenses",
        "This application uses the following open-source libraries:",
        ""
    ]

    for package in packages:
        info = get_package_license_info(package)
        if info and 'error' not in info:
            simple_content.append(f"• {package} v{info['version']} - {info['license']}")
        else:
            simple_content.append(f"• {package} - License information not available")

    simple_content.extend([
        "",
        "Complete license texts are available in THIRD_PARTY_LICENSES.txt",
        "or from the respective project websites."
    ])

    with open('../../THIRD_PARTY_LICENSES_SIMPLE.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(simple_content))


if __name__ == "__main__":
    print("Generating third-party license files...")
    generate_third_party_licenses()
    print("Done!")