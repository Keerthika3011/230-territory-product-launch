FROM public.ecr.aws/lambda/python:3.12
RUN pip install ortools openpyxl --target "${LAMBDA_TASK_ROOT}"
COPY solver.py ${LAMBDA_TASK_ROOT}
CMD ["solver.lambda_handler"]
