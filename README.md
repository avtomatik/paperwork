# Paperwork

Paperwork is a Python-based utility designed to streamline and automate various business processes. It offers a set of tools and scripts that facilitate tasks such as data transformation, document generation, and process automation, aiming to enhance efficiency and reduce manual intervention.

## Features

- Data Transformation: Automate the process of transforming and formatting data to meet specific requirements.
- Document Generation: Create and manage documents dynamically based on predefined templates and data inputs.
- Process Automation: Implement workflows that automate repetitive tasks, ensuring consistency and saving time.

## Installation

To get started with Paperwork, clone the repository and install the necessary dependencies:

```
git clone https://github.com/avtomatik/paperwork.git
cd paperwork
pip install --no-cache-dir -r requirements.txt
```

## Usage

### Main Script

The primary script in this repository is designed to process data and generate outputs based on specific configurations. To run the script:

```
python main.py --config config.yaml
```

Ensure that you have a config.yaml file with the appropriate settings.

### Configuration

The configuration file (config.yaml) should include the following parameters:

- source: Path to the source data file.
- data.file_name: Name of the data file.
- num: Number of records to process.

An example configuration:

```
source: /path/to/data
data:
  file_name: data.xlsx
num: 100
```

### Functions

- get_paths(): Reads a text file (PATHS.txt) and returns a tuple of paths. Each line in the file represents a separate path.

```
from pathlib import Path

def get_paths() -> tuple[Path]:
    paths_file = Path('../paths.txt')
    with paths_file.open(mode='r', encoding='utf-8') as f:
        return tuple(Path(line.strip()) for line in f if line.strip())
```

## Contributing

Contributions are welcome! Please fork the repository, create a new branch, and submit a pull request with your proposed changes.

## License

This project is licensed under the MIT License - see the [LICENSE.md](LICENSE.md) file for details.

---
