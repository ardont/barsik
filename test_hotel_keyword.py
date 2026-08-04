from engine.normalizer import classify_service

print("Before adding keyword:")
print("Размещение в гостинице:", classify_service("Размещение в гостинице"))
print("размещение 11111111111:", classify_service("размещение 11111111111"))

# Now let's test adding 'размещение'
import config
config.SERVICE_CLASSIFICATION["Hotel"].append("размещение")

print("\nAfter adding keyword 'размещение':")
print("Размещение в гостинице:", classify_service("Размещение в гостинице"))
print("размещение 11111111111:", classify_service("размещение 11111111111"))
