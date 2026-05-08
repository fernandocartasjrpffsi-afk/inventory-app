# Inventory App Version 1.2.2 - AI Assistant Accuracy Fix

This version improves the offline AI Assistant accuracy.

## Improved

- Better intent detection
- Better product keyword matching
- Better supplier/category/batch filtering
- Prevents overly narrow product matching
- Shows detected intent and filters
- More accurate status filters:
  - Available = stock > 0 and not expired
  - Expiring = stock > 0 and Expiring Soon/Urgent Expiry
  - Expired with stock = expired and stock > 0
  - Out of stock = stock <= 0

## Removed

No separate Management Summary or FEFO Generator buttons. The page keeps only the AI Assistant question box.

## Run

```bash
python -m streamlit run app.py
```
