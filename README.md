# Deep Learning with PyTorch — MLP, CNN, RNN & Seq2Seq

Comprehensive Deep Learning course project (EMSI, 2025-2026) covering the three canonical neural architectures — fully implemented, trained, and analyzed with PyTorch.

**Stack**: Python · PyTorch · Jupyter · NumPy · Pandas · Matplotlib

---

## Overview

Three self-contained notebooks, each going from first principles (manual math, hand-built layers) to a working trained model, with a critical experimental study.

| Part | Architecture | Dataset | Notebook |
|---|---|---|---|
| **I** | MLP (Multi-Layer Perceptron) | Adult Income (UCI) | [`notebooks/partie1_mlp.ipynb`](notebooks/partie1_mlp.ipynb) |
| **II** | CNN (Convolutional Neural Network) | CIFAR-10 | [`notebooks/partie2_cnn_cifar10.ipynb`](notebooks/partie2_cnn_cifar10.ipynb) |
| **III** | RNN / LSTM / GRU + Seq2Seq | The Time Machine, Tatoeba fra-eng | [`notebooks/partie3_rnn_seq2seq.ipynb`](notebooks/partie3_rnn_seq2seq.ipynb) |

Final report: [`Rapport_Deep_Learning.pdf`](Rapport_Deep_Learning.pdf)

---

## Part I — MLP on Adult Income

PyTorch fundamentals through an end-to-end binary classification task (predicting whether income exceeds \$50K/year from census data).

**Content**
- Core PyTorch concepts (tensors, autograd, `nn.Module`)
- Two equivalent MLP implementations: `nn.Sequential` vs custom `nn.Module` class
- Parameter inspection and initialization strategies
- Training loop on CPU/GPU
- Experimental comparison of configurations
- Best model save/load, final evaluation
- Critical analysis and synthesis

67 cells. Reproducible on CPU in a few minutes.

## Part II — CNN on CIFAR-10

From manual cross-correlation computations to a tuned CNN reaching **83.9% test accuracy**.

**Content**
- Manual computations: cross-correlation, output dimensions, padding, stride
- From-scratch convolution and pooling implementations validated against `torch.nn`
- Baseline CNN inspired by LeNet-5
- Parametrized improved variant (deeper 64→128→256, BatchNorm, dropout, weight decay, cosine scheduler, data augmentation)
- Experimental study: effect of padding, stride, pooling, number of filters, 1×1 convolutions
- Feature maps visualization
- MLP vs CNN comparison on the same dataset

**Results**

| Model | Test accuracy |
|---|---:|
| Baseline CNN (LeNet-5 style) | ~70% |
| **Improved CNN** (deeper + augmentation + BN/dropout) | **83.9%** (+14.3 pts) |

39 cells. The CIFAR-10 training needs a GPU to run in reasonable time.

## Part III — RNN, LSTM, GRU & Seq2Seq

### Volet A — Character-level language modeling (The Time Machine)
- Probabilistic framing of language models and perplexity
- Custom `Vocab` class and sequential mini-batching
- Hand-built RNN cell validated against `nn.RNN`
- Unified RNN / LSTM / GRU implementation
- Backpropagation through time (BPTT) and gradient clipping demonstration

**Results**

| Model | Final perplexity |
|---|---:|
| Vanilla RNN | 3.46 |
| GRU | 2.99 |
| **LSTM** | **2.78** |

Without gradient clipping, training diverges (perplexity > 5000). With clipping, all three converge and generate coherent text.

### Volet B — Seq2Seq translation (English → French)
- Encoder–decoder GRU architecture with teacher forcing
- Masked cross-entropy loss
- Greedy decoding **and** beam search
- From-scratch BLEU implementation

**Results**

| Decoding | BLEU (train) | BLEU (test) |
|---|---:|---:|
| Greedy | ~0.40 | ~0.03 |
| Beam search | ~0.43 | ~0.03 |

The train/test gap illustrates the overfitting of a small seq2seq without attention — motivating the attention mechanism.

42 cells.

---

## Project Structure

```
.
├── notebooks/
│   ├── partie1_mlp.ipynb          # Part I — MLP on Adult Income
│   ├── partie2_cnn_cifar10.ipynb  # Part II — CNN on CIFAR-10
│   └── partie3_rnn_seq2seq.ipynb  # Part III — RNN / LSTM / GRU / Seq2Seq
├── report_assets/                 # Figures embedded in the final report
├── Rapport_Deep_Learning.pdf      # Final written report
├── _build_report.py               # Script used to assemble the PDF report
└── README.md
```

## Reproducing Results

Each notebook is self-contained and ends with a verification cell. To re-run any part from the command line:

```bash
cd notebooks
jupyter nbconvert --to notebook --execute partie1_mlp.ipynb \
    --inplace --ExecutePreprocessor.timeout=1200

jupyter nbconvert --to notebook --execute partie2_cnn_cifar10.ipynb \
    --inplace --ExecutePreprocessor.timeout=7200

jupyter nbconvert --to notebook --execute partie3_rnn_seq2seq.ipynb \
    --inplace --ExecutePreprocessor.timeout=7200
```

Datasets are downloaded automatically by `torchvision` (CIFAR-10) or fetched from their canonical sources (Adult Income from UCI, Time Machine from Project Gutenberg, Tatoeba fra-eng).

## Requirements

- Python 3.10+
- PyTorch 2.x
- torchvision
- numpy, pandas, matplotlib, scikit-learn
- jupyter

A GPU is recommended for Parts II and III (training on CPU is possible but slow).

## Author

**Anass Habachi** — Student engineer in AI & Data Science (EMSI Casablanca).
Looking for a 6-month end-of-studies internship (PFE) in Data Engineering / AI starting February 2027.
